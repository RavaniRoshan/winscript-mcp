# File Triage Agent

**Goal:** An AI agent scans a messy Downloads folder, reads the contents of unidentified files, categorizes them based on their text, and moves them to the correct organized folders. 

This automation is completely mundane but saves users hours of weekly admin work.

### The Prompt
> "Scan the `C:/Downloads/To_Sort` folder. For every text file you find, read its contents. If the text looks like an invoice, move it to `C:/Documents/Invoices`. If it's a personal receipt, move it to `C:/Documents/Receipts`. Delete any files that are just empty or contain gibberish."

### WinScript Tool Execution Sequence

1. **`list_dir`**
   - **Arguments:** `{"path": "C:/Downloads/To_Sort"}`
   - *Agent receives the directory listing and identifies `[FILE] unknown_01.txt`, `[FILE] scan_02.txt`, etc.*

2. **`read_file_text`** *(Iterative)*
   - **Arguments:** `{"path": "C:/Downloads/To_Sort/unknown_01.txt"}`
   - *Agent reads the file. LLM analyzes the string. Identifies "Invoice #492" inside.*

3. **`move_file`**
   - **Arguments:** `{"src": "C:/Downloads/To_Sort/unknown_01.txt", "dst": "C:/Documents/Invoices/unknown_01.txt"}`

4. **`read_file_text`**
   - **Arguments:** `{"path": "C:/Downloads/To_Sort/scan_02.txt"}`
   - *Agent reads the file. LLM identifies empty lines or random characters.*

5. **`delete_file`**
   - **Arguments:** `{"path": "C:/Downloads/To_Sort/scan_02.txt"}`

### Why this works
- The MCP server's file system tools abstract away standard Python `pathlib` and `shutil` handling securely.
- LLM text comprehension replaces traditional fragile Regex parsers for categorization.