# Screenshot + Vision Fallback Loop

**Goal:** An AI agent attempts to automate an old legacy system or an Electron app (like Discord or a custom corporate tool) that has a completely broken accessibility tree. 

When standard UI elements can't be found via name or ID, the agent takes a screenshot, passes it to the Claude Vision model to visually locate the exact pixel coordinates of the target button, and executes a raw click. This is the closest thing to real Computer Use on Windows.

### The Prompt
> "Open the 'Legacy Inventory' app. Wait for it to load. We need to click the green 'Submit Batch' button, but the app's UI tree is broken. Please take a screenshot, visually locate the button, and use the coordinate click tool to click it."

### WinScript Tool Execution Sequence

1. **`open_app`**
   - **Arguments:** `{"name": "Legacy Inventory"}`

2. **`wait_for_window`**
   - **Arguments:** `{"title_hint": "Legacy Inventory", "timeout_seconds": 10}`
   - *Ensures the app is actually rendered on screen before proceeding.*

3. **`take_screenshot`**
   - **Arguments:** `{}` *(or a specific region if known)*
   - *Agent receives a `data:image/png;base64,...` string directly inline via MCP.*

4. **Vision Analysis (Internal LLM step)**
   - *The agent feeds the base64 image into its Vision model.*
   - *The Vision model identifies the "Submit Batch" button at `X: 1450, Y: 890`.*

5. **`coordinate_click`**
   - **Arguments:** `{"x": 1450, "y": 890}`
   - *WinScript forces the mouse to exactly that pixel and sends a left click.*

6. **`take_screenshot`** *(Verification)*
   - **Arguments:** `{}`
   - *Agent takes another screenshot to verify the screen state changed to the success confirmation page.*

### Why this works
- `take_screenshot` uses `mss` for instantaneous, in-memory frame grabs, bypassing slow disk I/O.
- `coordinate_click` is the ultimate fallback, meaning no Windows application is truly "un-automatable."