#!/usr/bin/env python3

import os
import json
import click
import requests
from bs4 import BeautifulSoup
import html2text
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import re
import webbrowser
import time
import urllib.parse
import msal
import http.server
import socketserver
import threading
import urllib.parse
import hashlib
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed

# Port for the local redirect server
REDIRECT_PORT = 8400

class AuthHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        """Handle the OAuth redirect."""
        # Parse the authorization code from the URL
        query = urllib.parse.urlparse(self.path).query
        params = urllib.parse.parse_qs(query)
        
        if 'code' in params:
            # Store the code in the server instance
            self.server.auth_code = params['code'][0]
            # Send a success page
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b"Authentication successful! You can close this window.")
        else:
            # Send an error page
            self.send_response(400)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b"Authentication failed! Please try again.")

class OneNoteClient:
    def __init__(self, client_id: str):
        self.client_id = client_id
        self.access_token = None
        self.graph_url = "https://graph.microsoft.com/v1.0"
        
    def get_access_token(self) -> str:
        """Get Microsoft Graph API access token using interactive browser flow."""
        app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority="https://login.microsoftonline.com/consumers"
        )
        
        # Start a local server to receive the redirect
        with socketserver.TCPServer(("", REDIRECT_PORT), AuthHandler) as httpd:
            # Store the server instance so we can access it from the handler
            httpd.auth_code = None
            
            # Generate the auth URL
            auth_url = app.get_authorization_request_url(
                scopes=["Notes.Read"],
                redirect_uri=f"http://localhost:{REDIRECT_PORT}",
                state="state"  # You might want to generate a random state
            )
            
            # Open the browser for authentication
            click.echo("Opening browser for authentication...")
            webbrowser.open(auth_url)
            
            # Wait for the auth code
            while not httpd.auth_code:
                httpd.handle_request()
            
            # Exchange the auth code for a token
            result = app.acquire_token_by_authorization_code(
                code=httpd.auth_code,
                scopes=["Notes.Read"],
                redirect_uri=f"http://localhost:{REDIRECT_PORT}"
            )
        
        if "access_token" not in result:
            raise Exception(f"Failed to get access token: {result.get('error_description', 'Unknown error')}")
            
        return result["access_token"]
    
    def _make_request(self, url: str, headers: Dict = None, params: Dict = None) -> Dict:
        """Make a request to the Microsoft Graph API with retry logic for timeouts."""
        if headers is None:
            headers = {}
        if params is None:
            params = {}
        
        # Ensure we have an access token
        if not self.access_token:
            self.access_token = self.get_access_token()
        
        # Add authorization header
        headers["Authorization"] = f"Bearer {self.access_token}"
        
        # Add default headers
        headers.update({
            "Accept": "application/json",
            "Content-Type": "application/json"
        })
        
        # Construct full URL if it's a relative path
        if not url.startswith('http'):
            url = f"{self.graph_url}/{url.lstrip('/')}"
        
        max_retries = 3
        retry_delay = 2  # seconds
        
        for attempt in range(max_retries):
            try:
                response = requests.get(url, headers=headers, params=params)
                if response.status_code == 401 and attempt < max_retries - 1:
                    # Token might have expired, try to get a new one
                    click.echo("Token expired, refreshing...", err=True)
                    self.access_token = self.get_access_token()
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    continue
                if response.status_code == 504 and attempt < max_retries - 1:
                    click.echo(f"Request timed out (504), retrying in {retry_delay} seconds... (Attempt {attempt + 1}/{max_retries})", err=True)
                    time.sleep(retry_delay)
                    retry_delay *= 2
                    continue
                response.raise_for_status()
                return response.json()
            except requests.exceptions.RequestException as e:
                if attempt < max_retries - 1:
                    click.echo(f"Request failed: {str(e)}, retrying in {retry_delay} seconds... (Attempt {attempt + 1}/{max_retries})", err=True)
                    time.sleep(retry_delay)
                    retry_delay *= 2
                    continue
                raise
        
        raise Exception("All retry attempts failed")
    
    def get_notebooks(self) -> List[Dict]:
        """Get list of available notebooks."""
        response = self._make_request("me/onenote/notebooks", params={})
        return response.get("value", [])
    
    def get_sections(self, notebook_id: str) -> List[Dict]:
        """Get sections in a notebook."""
        response = self._make_request(f"me/onenote/notebooks/{notebook_id}/sections", params={})
        return response.get("value", [])
    
    def get_pages(self, section_id: str) -> List[Dict]:
        """Get pages in a section, including nested pages."""
        all_pages = []
        page_count = 0
        skip = 0
        page_size = 100
        
        click.echo("Fetching pages from OneNote...")
        
        # Initial request parameters - only get what we need
        params = {
            "pagelevel": "true",
            "$select": "id,title,level,order",  # Only get the fields we need
            "$top": str(page_size),  # Get more pages per request
            "$count": "true"  # Include total count in response
        }
        
        headers = {"ConsistencyLevel": "eventual"}  # Required for $count
        
        # Keep fetching pages until we have them all
        while True:
            # Update skip parameter for pagination
            params["$skip"] = str(skip)
            
            click.echo(f"Fetching pages {skip + 1} to {skip + page_size}...")
            response = self._make_request(
                f"me/onenote/sections/{section_id}/pages",
                headers=headers,
                params=params
            )
            
            # Debug: Print response metadata
            click.echo(f"Response metadata: {json.dumps({k: v for k, v in response.items() if k.startswith('@odata')}, indent=2)}")
            
            # Add pages from this response
            pages = response.get("value", [])
            if not pages:  # No more pages to process
                break
                
            all_pages.extend(pages)
            page_count += len(pages)
            
            # Get total count if available
            total_count = response.get("@odata.count")
            if total_count:
                click.echo(f"Retrieved {len(pages)} pages (total: {page_count} of {total_count})")
            else:
                click.echo(f"Retrieved {len(pages)} pages (total: {page_count})")
            
            # Check if we need to fetch more pages
            if total_count and page_count < total_count:
                skip += page_size
                click.echo(f"More pages available, will fetch next batch starting at {skip + 1}...")
            else:
                click.echo("No more pages to fetch.")
                break
        
        if not all_pages:
            click.echo("No pages found in this section.")
            return []
            
        click.echo(f"\nProcessing {len(all_pages)} pages to build hierarchy...")
        
        # Sort pages by their order property first, then by level
        all_pages.sort(key=lambda x: (x.get('order', 0), x.get('level', 0)))
        
        # Create a map of page IDs to their children
        page_map = {}
        root_pages = []
        
        # Process pages in order
        for page in all_pages:
            page_id = page['id']
            page['children'] = []
            page_map[page_id] = page
            
            level = page.get('level', 0)
            if level == 0:
                # This is a root page
                root_pages.append(page)
            else:
                # Find the parent page by looking at the previous pages
                # The parent is the last page we processed with a level one less than current
                parent = None
                for prev_page in reversed(list(page_map.values())):
                    if prev_page.get('level', 0) == level - 1:
                        parent = prev_page
                        break
                
                if parent:
                    parent['children'].append(page)
                else:
                    # If we can't find a parent, treat it as a root page
                    root_pages.append(page)
        
        click.echo(f"Found {len(root_pages)} root pages with their children")
        return root_pages
    
    def get_page_content(self, page_id: str) -> str:
        """Get the content of a page."""
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Prefer": "outlook.body-content-type=text"  # Request text content directly
        }
        response = requests.get(
            f"{self.graph_url}/me/onenote/pages/{page_id}/content",
            headers=headers
        )
        
        if response.status_code != 200:
            raise Exception(f"Failed to get page content: {response.status_code} - {response.text}")
            
        return response.text

class OneNoteToMarkdown:
    def __init__(self, client: OneNoteClient):
        self.client = client
        self.h2t = html2text.HTML2Text()
        self.h2t.ignore_links = False
        self.h2t.ignore_images = False
        self.h2t.ignore_tables = False
        self.image_counter = 0  # For generating unique image filenames
        
    def sanitize_filename(self, filename: str) -> str:
        """Sanitize filename to be safe for filesystem while preserving readability."""
        if not filename:
            return 'untitled'
        
        # Replace slashes with dashes
        filename = filename.replace('/', '-').replace('\\', '-')
        
        # Remove or replace unsafe characters for filesystems
        # Keep: letters, numbers, spaces, hyphens, underscores, dots, parentheses, brackets
        # Remove: control characters, and other potentially problematic chars
        import re
        # Remove control characters and other unsafe characters
        filename = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', filename)
        
        # Replace multiple consecutive dashes or spaces with single ones
        filename = re.sub(r'-+', '-', filename)
        filename = re.sub(r' +', ' ', filename)
        
        # Trim leading and trailing spaces, dashes, and dots
        filename = filename.strip(' .-')
        
        # Ensure the filename is not empty after sanitization
        if not filename:
            return 'untitled'
        
        return filename

    def sanitize_image_filename(self, filename: str) -> str:
        """Sanitize image filename to be safe for filesystem with lowercase and dashes."""
        if not filename:
            return 'untitled'
        
        # Replace slashes with dashes
        filename = filename.replace('/', '-').replace('\\', '-')
        
        # Convert to lowercase
        filename = filename.lower()
        
        # Replace spaces and special characters with dashes
        import re
        # Replace spaces and special characters with dashes
        filename = re.sub(r'[^a-z0-9.-]', '-', filename)
        
        # Remove control characters
        filename = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', filename)
        
        # Replace multiple consecutive dashes with single ones
        filename = re.sub(r'-+', '-', filename)
        
        # Trim leading and trailing dashes and dots
        filename = filename.strip('.-')
        
        # Ensure the filename is not empty after sanitization
        if not filename:
            return 'untitled'
        
        return filename

    def download_image(self, image_url: str, images_dir: Path, page_title: str, headers: Optional[Dict[str, str]] = None) -> Optional[str]:
        """Download an image and return its local path relative to the markdown file."""
        try:
            # Generate a unique filename using page name, URL hash and counter
            page_name = self.sanitize_image_filename(page_title)
            url_hash = hashlib.md5(image_url.encode()).hexdigest()[:8]
            self.image_counter += 1
            
            # Try to get extension from URL or content type
            ext = os.path.splitext(urlparse(image_url).path)[1]
            if not ext:
                # Make a HEAD request to get content type if we have headers
                if headers:
                    response = requests.head(image_url, headers=headers)
                    content_type = response.headers.get('content-type', '')
                    if 'image/' in content_type:
                        ext = '.' + content_type.split('/')[-1]
            if not ext:
                ext = '.png'  # Default to .png if we can't determine the type
            if not ext.startswith('.'):
                ext = '.' + ext
                
            # Sanitize the extension as well
            ext = self.sanitize_filename(ext.lstrip('.'))
            if not ext.startswith('.'):
                ext = '.' + ext
                
            local_filename = f"{page_name}_image_{url_hash}_{self.image_counter}{ext}"
            local_path = images_dir / local_filename
            
            # Download the image
            response = requests.get(image_url, stream=True, headers=headers)
            response.raise_for_status()
            
            # Save the image
            with open(local_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            # Return the relative path for use in markdown
            return f"images/{local_filename}"
        except Exception as e:
            click.echo(f"Warning: Failed to download image {image_url}: {str(e)}", err=True)
            return None
    
    def _clean_machine_generated_alt_text(self, soup: BeautifulSoup) -> None:
        """Remove machine-generated alt text from images."""
        for img in soup.find_all('img'):
            if img.get('alt', '').startswith('Machine generated alternative text:'):
                img['alt'] = ''

    def _convert_bold_spans(self, soup: BeautifulSoup) -> None:
        """Wrap contents of spans with font-weight:bold in b elements."""
        for span in soup.find_all('span'):
            style = span.get('style', '')
            if 'font-weight:bold' in style or 'font-weight: 700' in style:
                # Create a new b element
                b = soup.new_tag('b')
                # Move all contents of the span to the b element
                b.extend(span.contents)
                # Clear the span and append the b element
                span.clear()
                span.append(b)

    def convert_page_to_markdown(self, html_content: str, images_dir: Path, page_title: str, is_child_page: bool = False) -> str:
        """Convert OneNote HTML content to Markdown and download images."""
        # Parse HTML
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Clean up machine-generated alt text
        self._clean_machine_generated_alt_text(soup)
        
        # Convert bold spans to b elements
        self._convert_bold_spans(soup)
        
        # Handle OneNote-specific elements and download images
        for img in soup.find_all('img'):
            # Get the image URL - prefer data-fullres-src for OneNote images
            img_url = None
            if 'data-fullres-src' in img.attrs:
                img_url = img['data-fullres-src']
            elif 'src' in img.attrs:
                img_url = img['src']
            
            if not img_url:
                continue
                
            # For Microsoft Graph API image URLs, we need to include the access token
            if 'graph.microsoft.com' in img_url and self.client.access_token:
                # Add authorization header to the request
                headers = {"Authorization": f"Bearer {self.client.access_token}"}
            else:
                headers = None
                
            try:
                # Download the image and get local path
                local_path = self.download_image(img_url, images_dir, page_title, headers)
                if local_path:
                    # For child pages, adjust the image path to be relative to the parent directory
                    if is_child_page:
                        local_path = f"../{local_path}"
                    img['src'] = local_path
                    # Remove data attributes as they're no longer needed
                    for attr in ['data-src-type', 'data-fullres-src', 'data-fullres-src-type']:
                        if attr in img.attrs:
                            del img[attr]
            except Exception as e:
                click.echo(f"Warning: Failed to process image {img_url}: {str(e)}", err=True)
                continue
        
        # Simplify links where text and URL are identical
        for a in soup.find_all('a'):
            href = a.get('href', '')
            if href and href == a.text.strip():
                # Replace the link with just the URL text
                a.replace_with(href)
        
        # Convert to Markdown
        markdown = self.h2t.handle(str(soup))
        
        # Clean up the markdown
        lines = [line.strip() for line in markdown.splitlines()]
        cleaned = []
        prev_blank = False
        for line in lines:
            if line == '':
                if not prev_blank:
                    cleaned.append('')
                prev_blank = True
            else:
                cleaned.append(line)
                prev_blank = False
        # Remove leading/trailing blank lines
        while cleaned and cleaned[0] == '':
            cleaned.pop(0)
        while cleaned and cleaned[-1] == '':
            cleaned.pop()
        return '\n'.join(cleaned)
    
    def process_page(self, page: Dict, output_path: Path, parent_dir: Path = None) -> None:
        """Process a single page."""
        page_title = page["title"]
        sanitized_title = self.sanitize_filename(page_title)
        has_children = bool(page.get('children', []))
        
        # Determine the directory for this page and its files
        if parent_dir:
            # This is a child page, save directly in parent's directory
            page_dir = parent_dir
            # Ensure parent directory exists
            page_dir.mkdir(parents=True, exist_ok=True)
            # Save files in the parent's directory
            markdown_path = page_dir / f"{sanitized_title}.md"
            is_child_page = True
        else:
            # This is a top-level page
            if has_children:
                # Only create a directory if the page has children
                page_dir = output_path / sanitized_title
                page_dir.mkdir(parents=True, exist_ok=True)
                markdown_path = output_path / f"{sanitized_title}.md"
            else:
                # No children, save directly in output directory
                page_dir = output_path
                markdown_path = output_path / f"{sanitized_title}.md"
            is_child_page = False
        
        # Use a single root-level images directory for all pages in this section
        images_dir = output_path / "images"
        images_dir.mkdir(parents=True, exist_ok=True)
        
        # Reset image counter for this page
        self.image_counter = 0
        
        # Get page content
        html_content = self.client.get_page_content(page["id"])
        
        # Convert to Markdown
        markdown_content = self.convert_page_to_markdown(html_content, images_dir, page_title, is_child_page)
        
        # Ensure the directory exists before writing the file
        markdown_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Save markdown file
        with open(markdown_path, 'w', encoding='utf-8') as f:
            f.write(f"# {page_title}\n\n")
            f.write(markdown_content)
        
        click.echo(f"Converted: {page_title} -> {markdown_path}")
    
    def print_page_hierarchy(self, pages: List[Dict], level: int = 0) -> None:
        """Print the page hierarchy in a tree-like structure."""
        for page in pages:
            click.echo(f"{'  ' * level}• {page['title']}")
            self.print_page_hierarchy(page.get('children', []), level + 1)

    def download_and_convert(self, notebook_name: str, section_name: Optional[str], output_dir: str):
        """Download pages from a section or all sections and convert them to Markdown."""
        # Create base output directory
        base_output_path = Path(output_dir)
        base_output_path.mkdir(parents=True, exist_ok=True)
        
        click.echo(f"\nLooking for notebook '{notebook_name}'...")
        # Find notebook
        notebooks = self.client.get_notebooks()
        notebook = next((n for n in notebooks if n["displayName"] == notebook_name), None)
        if not notebook:
            raise Exception(f"Notebook '{notebook_name}' not found")
        click.echo(f"Found notebook: {notebook_name}")
        
        # Get all sections
        sections = self.client.get_sections(notebook["id"])
        if not sections:
            raise Exception(f"No sections found in notebook '{notebook_name}'")
            
        # Filter sections if section_name is provided
        if section_name:
            sections = [s for s in sections if s["displayName"] == section_name]
            if not sections:
                raise Exception(f"Section '{section_name}' not found in notebook '{notebook_name}'")
            click.echo(f"\nDownloading section: {section_name}")
        else:
            click.echo(f"\nDownloading all sections ({len(sections)} sections found):")
            for section in sections:
                click.echo(f"  • {section['displayName']}")
        
        # Process each section
        for section in sections:
            section_name = section["displayName"]
            click.echo(f"\nProcessing section: {section_name}")
            
            # Create section-specific output directory
            section_output_path = base_output_path / section_name
            section_output_path.mkdir(parents=True, exist_ok=True)
            
            # Create section-specific images directory
            images_dir = section_output_path / "images"
            images_dir.mkdir(parents=True, exist_ok=True)
            
            click.echo("\nStarting page download and conversion...")
            # Get and convert pages (including nested pages)
            pages = self.client.get_pages(section["id"])
            
            def count_pages(page):
                count = 1  # Count the current page
                for child in page.get('children', []):
                    count += count_pages(child)
                return count
            
            total_pages = sum(count_pages(page) for page in pages)
            
            click.echo("\nPage hierarchy:")
            self.print_page_hierarchy(pages)
            click.echo(f"\nConverting {total_pages} pages to Markdown...")
            
            # Collect all pages (including children) for parallel processing
            all_pages = []
            
            def collect_pages(page, parent_dir=None):
                all_pages.append((page, parent_dir))
                # For children, pass the parent's directory path
                if parent_dir is None:
                    # Top-level page - children will be in the page's directory
                    page_dir = section_output_path / self.sanitize_filename(page["title"])
                else:
                    # Child page - children will be in the parent's directory
                    page_dir = parent_dir
                
                for child in page.get('children', []):
                    collect_pages(child, page_dir)
            
            for page in pages:
                collect_pages(page)
            
            # Process all pages in parallel
            failed_pages = []
            
            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = []
                for page, parent_dir in all_pages:
                    future = executor.submit(self.process_page, page, section_output_path, parent_dir)
                    futures.append((future, page))
                
                # Wait for all pages to complete
                for future, page in futures:
                    try:
                        future.result()
                    except Exception as e:
                        error_msg = str(e)
                        click.echo(f"Error processing page '{page['title']}': {error_msg}", err=True)
                        failed_pages.append((page['title'], error_msg))
                        # Continue processing other pages even if one fails
            
            # Display summary of failed pages
            if failed_pages:
                click.echo(f"\n❌ Failed to process {len(failed_pages)} pages in section '{section_name}':")
                for page_title, error in failed_pages:
                    click.echo(f"  • {page_title}: {error}")
            else:
                click.echo(f"\n✅ Successfully processed all {total_pages} pages in section '{section_name}'")
            
            click.echo(f"\nCompleted conversion for section: {section_name}")
        
        click.echo("\nAll conversions completed successfully!")

@click.group()
def cli():
    """OneNote to Markdown converter CLI."""
    pass

@cli.command()
@click.option('--notebook', required=True, help='Name of the OneNote notebook')
@click.option('--section', help='Name of the section to download (if not specified, downloads all sections)')
@click.option('--output-dir', default='./output', help='Output directory for Markdown files')
@click.option('--client-id', default='8e1a6f85-d243-41ac-a6d3-4b7fd05ce004', help='Microsoft Graph API client ID (defaults to shared application)')
def download(notebook: str, section: Optional[str], output_dir: str, client_id: str):
    """Download and convert OneNote pages to Markdown format."""
    try:
        # Initialize the client with the provided client ID
        client = OneNoteClient(client_id)
        converter = OneNoteToMarkdown(client)
        converter.download_and_convert(notebook, section, output_dir)
    except Exception as e:
        click.echo(f"Error: {str(e)}", err=True)
        raise click.Abort()

if __name__ == '__main__':
    cli() 