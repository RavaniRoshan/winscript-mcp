# Excel → Email Pipeline

**Goal:** An AI agent reads a corporate sales report from Excel, analyzes the data, formats a summary, and emails it directly to the boss.

This is a high-value "wow" demo for corporate environments where manual data reporting takes up hours each week.

### The Prompt
> "Please read the Q3 Sales Report at `C:/Reports/Q3_Sales.xlsx`. Extract the data from `Sheet1`, range `A1:E20`. Summarize the top 3 performing regions and their total revenue. Format the summary nicely and email it to `boss@company.com` with the subject 'Q3 Sales Top Performers Summary'."

### WinScript Tool Execution Sequence

1. **`excel_read_range`**
   - **Arguments:** `{"filepath": "C:/Reports/Q3_Sales.xlsx", "sheet": "Sheet1", "start": "A1", "end": "E20"}`
   - *Agent analyzes the CSV-style string returned by the tool, calculates the top 3 regions, and formats a written summary.*

2. **`outlook_send_email`**
   - **Arguments:** 
     ```json
     {
       "to": "boss@company.com",
       "subject": "Q3 Sales Top Performers Summary",
       "body": "Here are the top 3 regions for Q3:\n\n1. North America - $1.2M\n2. Europe - $950K\n3. APAC - $800K\n\nLet me know if you need the full breakdown."
     }
     ```
   - *Agent returns the confirmation string: "Email sent to boss@company.com | Subject: Q3 Sales Top Performers Summary".*

### Why this works
- WinScript handles the messy COM bindings in the background.
- Hidden `Excel.Application` ensures the UI never flashes and distracts the user.
- Explicit cleanup guarantees no orphaned processes hold file locks on the network drives.