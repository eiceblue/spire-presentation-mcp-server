## What is Spire.Presentation MCP Server?

The Spire.Presentation MCP Server is a robust solution that empowers AI agents to work with Excel files using the Model Context Protocol (MCP). It is totally independent and doesn't require Microsoft Office to be installed on system. This tool enables AI agents to [add or delete](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Document-Operation/Python-Add-or-Delete-Slides-in-PowerPoint-Presentations.html), read, and [convert PowerPoint Presentation](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Conversion/Python-Convert-PowerPoint-Presentation-to-PDF.html) seamlessly

## Main Features: 

- Convert [PowerPoint to PDF](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Conversion/Python-Convert-PowerPoint-Presentation-to-PDF.html), [PowerPoint to HTML](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Conversion/Python-Convert-PowerPoint-to-HTML.html), [PowerPoint to Image](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Conversion/Python-Convert-PowerPoint-to-Images-PNG-JPG-BMP-SVG.html), and more with high fidelity.
- Create, modify, and manage PowerPoint presentation
- Manage and control presentation: [create smartart](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/SmartArt/Python-Create-Read-or-Delete-SmartArt-in-PowerPoint.html), [create table](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Table/Python-Create-or-Edit-Tables-in-PowerPoint-Presentations.html), and more.
- Manage shape
- [group and ungroup](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Image-and-Shapes/Python-Group-or-Ungroup-Shapes-in-PowerPoint.html)
- Add various chart types
- [create a pie chart ](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Chart/Python-Create-a-Pie-Chart-or-a-Doughnut-Chart-in-PowerPoint.html) 

## How to use Spire.Presentation MCP Server?

### Prerequisites

- Python 3.10 or higher

### Installation

1. Clone the repository:
```bash
git clone https://github.com/eiceblue/spire-presentation-mcp-server
cd spire-presentation-mcp-server
```

2. Install using uv:
```bash
uv pip install -e .
```
### Running the Server

Start the server (default port 8000):
```bash
uv run spire-presentation-mcp-server
```

Custom port (e.g., 8080):

```bash
# Bash/Linux/macOS
export FASTMCP_PORT=8080 && uv run spire-presentation-mcp-server

# Windows PowerShell
$env:FASTMCP_PORT = "8080"; uv run spire-presentation-mcp-server
```

## Integration with AI Tools

### Cursor IDE

1. Add this configuration to Cursor:
```json
{
  "mcpServers": {
    "ppt": {
      "url": "http://localhost:8000/sse",
      "env": {
        "PPT_FILES_PATH": "/path/to/ppt/files"
      }
    }
  }
}
```
2. The PowerPoint tools will be available through your AI assistant.

### Remote Hosting & Transport Protocols

This server uses Server-Sent Events (SSE) transport protocol. For different use cases:

1. **Using with Claude Desktop (requires stdio):**
   - Use [Supergateway](https://github.com/supercorp-ai/supergateway) to convert SSE to stdio

2. **Hosting Your MCP Server:**
   - [Remote MCP Server Guide](https://developers.cloudflare.com/agents/guides/remote-mcp-server/)

## Environment Variables

| Variable | Description | Default |
|--------|------|--------|
| `FASTMCP_PORT` | Server port | `8000` |
| `PPT_FILES_PATH` | Directory for Presentation files | `./ppt_files` |

## Available Tools

The server provides a comprehensive set of PowerPoint manipulation tools. Here are the main categories:

- **Basic Operations**: Add, or [delete slide](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Document-Operation/Python-Add-or-Delete-Slides-in-PowerPoint-Presentations.html).
- **Data Processing**: Create or Edit Tables in PowerPoint Presentations, [apply formulas](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Table/Python-Create-or-Edit-Tables-in-PowerPoint-Presentations.html)
- **Formatting**: Apply styles, [Save Shapes](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Image-and-Shapes/Python-Save-Shapes-as-Image-Files-in-PowerPoint-Presentations.html) as Image Files in PowerPoint Presentations
- **Advanced Features**: [Create charts](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Chart/Python-Create-a-Pie-Chart-or-a-Doughnut-Chart-in-PowerPoint.html), [Create table](https://www.e-iceblue.com/Tutorials/Python/Spire.Presentation-for-Python/Program-Guide/Table/Python-Create-or-Edit-Tables-in-PowerPoint-Presentations.html)
- **Conversion**: Convert PowerPoint to PDF, HTML, image, and more with high fidelity.

See [TOOLS.md](https://github.com/eiceblue/spire-presentation-mcp-server/blob/main/TOOLS.md) for complete documentation of all available tools.

## FAQ from Spire.Presentation MCP Server?

Q1. Can I use Spire.Presentation MCP Server for any directory?

Yes, Spire.Presentation MCP Serer works for any directory.

Q2. Is Spire.Presentation MCP Server free to use?

Yes, it is licensed under the MIT License, allowing free use and modification.

Q3. What programming languages does Spire.Presentation MCP Server support?

It is built with Python.

## License
MIT
