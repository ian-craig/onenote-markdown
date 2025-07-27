# OneNote to Markdown Converter

A command-line utility that converts OneNote notebooks to Markdown format, preserving page hierarchy, images, and formatting. Perfect for migrating your OneNote content to other Markdown-based systems or version control.

## Features

- 🔐 Secure authentication with Microsoft Graph API
- 📚 Downloads entire notebooks or specific sections
- 📂 Preserves page hierarchy and folder structure
- 🖼️ Handles images and attachments
- 🔄 Maintains page relationships (parent/child)
- 🎨 Converts OneNote formatting to clean Markdown
- 🔍 Supports pagination for large notebooks
- 🔁 Automatic retry for failed requests
- 📝 Detailed progress logging
- 📁 Organizes output by section with separate image directories

## Prerequisites

- Python 3.11 or higher
- A Microsoft personal account with OneNote notebooks

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/onenote-markdown.git
   cd onenote-markdown
   ```

2. Install dependencies using `uv` (recommended) or `pip`:

   Using `uv` (faster):
   ```bash
   # Install uv if you haven't already
   curl -LsSf https://astral.sh/uv/install.sh | sh
   
   # Install dependencies
   uv pip install -r requirements.txt
   ```

   Using `pip`:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Download a specific section:
```bash
uv run onenote_markdown.py download --notebook "My Notebook" --section "My Section" --output-dir ./output
```

Download all sections in a notebook:
```bash
uv run onenote_markdown.py download --notebook "My Notebook" --output-dir ./output
```

The application uses a pre-configured Microsoft Graph API application by default - no additional setup required!

### Command Line Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `--notebook` | Yes | Name of the OneNote notebook to download |
| `--section` | No | Name of the section to download (if not specified, downloads all sections) |
| `--output-dir` | No | Directory where Markdown files will be saved (default: `./output`) |
| `--client-id` | No | Microsoft Graph API client ID (uses default shared application if not specified) |

### Authentication Flow

1. When you run the script, it will:
   - Open your default web browser
   - Prompt you to sign in to your Microsoft account
   - Request permission to access your OneNote notebooks
   - After authentication, you can close the browser window
   - The script will continue downloading and converting your notes

### Output Structure

The script creates a directory structure based on your notebook sections:

```
output/
├── Section1/              # Each section gets its own directory
│   ├── images/           # Section-specific images
│   ├── Page1.md         # Top-level pages
│   ├── PageWithChildren.md
│   └── PageWithChildren/
│       ├── Child1.md    # Child pages
│       └── Child2.md
│
├── Section2/              # Another section
│   ├── images/           # Separate images directory
│   ├── Page1.md
│   └── ...
│
└── Section3/              # And so on...
    ├── images/
    └── ...
```

Each section directory contains:
- A separate `images` directory for that section's images
- Markdown files for all pages in that section
- Subdirectories for pages with children
- Proper relative paths for images and links

## Advanced: Using Your Own Microsoft Graph Application

For additional security or customization, you can create your own Microsoft Graph application:

1. Register a new application:
   - Visit [Azure Portal - App Registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
   - Click "New registration"
   - Name: "OneNote Markdown Converter"
   - Supported account types: "Personal Microsoft accounts only"
   - Click "Register"
   - Copy the Application (client) ID

2. Configure the application:
   - In your registered app, go to "Authentication"
   - Under "Platform configurations", click "Add a platform"
   - Choose "Mobile and desktop applications"
   - Add the redirect URI: `http://localhost:8400`
   - Under "Default client type", check "Treat application as a public client"
   - Click "Configure"
   - Go to "API permissions"
   - Click "Add a permission"
   - Choose "Microsoft Graph"
   - Select "Delegated permissions"
   - Search for and add "Notes.Read"
   - Click "Grant admin consent" (if available)

3. Use your custom application:
   ```bash
   uv run onenote_markdown.py download --notebook "My Notebook" --client-id "your-client-id-here"
   ```

## Troubleshooting

### Common Issues

1. **Authentication Failed**
   - Ensure you provided the correct `--client-id`
   - Verify the application is properly configured in Azure Portal
   - Check that port 8400 is available

2. **Notebook/Section Not Found**
   - Verify the notebook and section names exactly match your OneNote
   - Names are case-sensitive
   - If downloading all sections, ensure the notebook exists

3. **Image Download Issues**
   - Check your internet connection
   - Verify you have write permissions in the output directory
   - Each section has its own images directory

### Getting Help

If you encounter any issues:
1. Check the error message for specific details
2. Ensure all prerequisites are met
3. Verify your Microsoft Graph API setup
4. Open an issue on GitHub with:
   - The exact command you ran
   - The complete error message
   - Your Python version
   - Your operating system

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details. 