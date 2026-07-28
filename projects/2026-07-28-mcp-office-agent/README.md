# MCP Office Agent

An enterprise-ready **Model Context Protocol (MCP)** server that empowers AI coding assistants (like Cursor, Windsurf, or Claude Desktop) to programmatically read, write, edit, and inspect Microsoft Office documents (`.docx`, `.xlsx`, and `.pptx`) using pure local Python libraries.

---

## 🌟 Overview

Large Language Models (LLMs) and coding agents excels at working with plain text, JSON, and Markdown. However, interacting with binary enterprise formats like Microsoft Word, Excel, and PowerPoint has historically required complex format conversions or external SaaS dependencies.

`mcp-office-agent` acts as a local-first bridge. Built on the Model Context Protocol, it exposes tools that allow LLMs to directly read and write Office documents programmatically. The server formats incoming documents into clean Markdown so the LLM can understand them, and translates JSON-based block schemas back into beautifully styled Office files.

---

## ⚡ Features

### 📝 Microsoft Word (`.docx`)
*   **Structure-Aware Reading:** Converts Word paragraphs, multi-level lists, headers, and grid tables into clean, structured Markdown.
*   **Block-Based Generation:** Generates styled `.docx` files using a simple JSON block format (supporting headings, paragraphs, styled list items, and tabular grids).

### 📊 Microsoft Excel (`.xlsx`)
*   **Formula & Style Preservation:** Supports reading formulas as code blocks or evaluating cached spreadsheet values.
*   **Automatic Formatting:** Generates sheets with a premium aesthetic: automated column width adjustments, grid line visibility, and stylized header rows (Classic Navy/Slate Blue headers with white text).

### 📉 Microsoft PowerPoint (`.pptx`)
*   **Outline Extraction:** Reads slide structure, text boxes, and slide speaker notes.
*   **Layout Templates:** Creates slide decks using standard layouts (`title_slide`, `title_and_content`, `title_only`, and `blank`), and writes speaker notes for presenters.

### 🔍 Metadata Inspection
*   **Lightweight Peeking:** Interrogate files to check slide outlines, sheet names, or paragraph/table counts before doing a full read.

---

## 🏗️ Architecture

```
                 +-----------------------------------+
                 |           MCP Client              |
                 |  (Claude Desktop, Cursor, etc.)   |
                 +-----------------+-----------------+
                                   |
                          JSON-RPC over stdio
                                   |
                                   v
                 +-----------------+-----------------+
                 |        FastMCP Server             |
                 |       (mcp-office-agent)          |
                 +--------+--------+--------+--------+
                          |        |        |
         +----------------+        |        +---------------+
         |                         |                        |
         v                         v                        v
+--------+--------+       +--------+--------+      +--------+--------+
|   word.py       |       |   excel.py      |      |  powerpoint.py  |
|  (python-docx)  |       |  (openpyxl)     |      |  (python-pptx)  |
+--------+--------+       +--------+--------+      +--------+--------+
         |                         |                        |
         v                         v                        v
   Word Documents           Excel Sheets              Slide Decks
      (.docx)                  (.xlsx)                  (.pptx)
```

---

## 🛠️ Tech Stack
*   **Language:** Python >= 3.10
*   **MCP Framework:** `mcp` SDK (FastMCP)
*   **Parsing Engines:** `python-docx`, `openpyxl`, `python-pptx`
*   **Test Suite:** `pytest`

---

## 📥 Installation

1.  Clone the repository or download the source files.
2.  Navigate to the project root and install the dependencies:
    ```bash
    pip install -e .
    ```

---

## 🚀 Setup & Integration

### Claude Desktop Configuration
To connect this agent to your Claude Desktop application, add the following configuration to your `claude_desktop_config.json` file (typically located at `%APPDATA%/Claude/claude_desktop_config.json` on Windows):

```json
{
  "mcpServers": {
    "office-agent": {
      "command": "python",
      "args": [
        "-m",
        "mcp_office_agent.server"
      ],
      "env": {
        "PYTHONPATH": "C:/Users/Gurteg/.gemini/antigravity/scratch/mcp-office-agent/src"
      }
    }
  }
}
```

### Cursor Integration
1. Open Cursor Settings.
2. Navigate to **Features** > **MCP**.
3. Click **+ Add New MCP Server**.
4. Set name to `office-agent`, Type to `command`, and Command to:
   ```bash
   python -m mcp_office_agent.server
   ```
5. Ensure the Python environment where you installed the dependencies is active, or use an absolute path to the Python executable.

---

## 📖 Usage Guide & Tool Schemas

### Available Tools

#### 1. `inspect_office_file`
*   **Purpose:** Reads basic structure/metadata.
*   **Input Schema:** `{"path": "string"}`

#### 2. `read_word_document` / `write_word_document`
*   **Read:** Converts docx text and tables into markdown.
*   **Write:** Writes a list of structured blocks. Example block schema:
    ```json
    [
      {"type": "heading", "level": 1, "text": "Report Header"},
      {"type": "paragraph", "text": "This is introductory text."},
      {"type": "list_item", "style": "List Bullet", "text": "Bullet Item"},
      {"type": "table", "headers": ["Name", "Score"], "rows": [["Alice", "95"], ["Bob", "88"]]}
    ]
    ```

#### 3. `read_excel_spreadsheet` / `write_excel_spreadsheet`
*   **Read:** Converts worksheets into side-by-side markdown tables.
*   **Write:** Maps sheets to 2D lists of rows. Supports equations:
    ```json
    {
      "Sales": [
        ["Month", "Revenue"],
        ["Jan", 1200],
        ["Feb", 1500],
        ["Total", "=SUM(B2:B3)"]
      ]
    }
    ```

#### 4. `read_powerpoint_presentation` / `write_powerpoint_presentation`
*   **Read:** Converts slides, shapes, and speaker notes to markdown.
*   **Write:** Generates presentations using standard slide layout configurations:
    ```json
    [
      {
        "layout": "title_slide",
        "title": "Annual Review",
        "subtitle": "Q4 Performance Summary"
      },
      {
        "layout": "title_and_content",
        "title": "Key Accomplishments",
        "content": ["Launched new web application", "Achieved 20% growth"],
        "notes": "Be sure to highlight the developer team performance here."
      }
    ]
    ```

---

## 🧪 Testing

To run the automated pytest suite:
```bash
pytest src/mcp_office_agent/tests
```

---

## 🔮 Future Improvements
*   **Advanced Styling:** Add support for changing font families, colors, and margins on Word paragraphs.
*   **Chart Generation:** Programmatically generate and embed PowerPoint/Excel native charts from raw data.
*   **Image Insertion:** Support inserting images into Word documents and PowerPoint slides.

---

## 📄 License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
