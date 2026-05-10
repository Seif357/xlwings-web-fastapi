# Excel xlwings Setup Guide

## 1. Set up GitHub Codespace
First, create a new GitHub Codespace. Once it's ready, open the terminal and execute the following commands to set up your Python environment:

```bash
# 1. Create the virtual environment
python3 -m venv venv

# 2. Activate it (Crucial step!)
source venv/bin/activate

# 3. Upgrade core tools
pip install --upgrade pip setuptools wheel

# 4. Install the LATEST versions (to ensure compatibility with Python 3.12)
pip install fastapi pandas cryptography httpx uvicorn

# 5. Run your app (ensure you are using uvicorn for FastAPI)
python app/main.py
```

## 2. Get and Set xlwings License Key
Visit [https://www.xlwings.org/trial](https://www.xlwings.org/trial) to get a trial key. Then, in your Codespace terminal, run:

```bash
export XLWINGS_LICENSE_KEY="your-trial-key-goes-here"
```

## 3. Set up Excel on the Web and Script Lab
1. Open **Excel on the web**.
2. Go to **Insert > Add-ins** (or Home > Add-ins) and search for **Script Lab**, then add it.
3. Open the **Script Lab** task pane and create a new snippet.
4. Replace the contents of each tab with the code provided below.

### Script Lab: Libraries
Add the following URLs to the **Libraries** tab:
```text
https://appsforoffice.microsoft.com/lib/1/hosted/office.js
https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts
https://cdn.jsdelivr.net/gh/xlwings/xlwings@0.30.1/xlwingsjs/dist/xlwings.min.js
```

### Script Lab: Script (JavaScript)
Paste the following code into the **Script** tab:

```javascript
declare let xlwings: any;

document.getElementById("run").addEventListener("click", () => tryCatch(run));

async function run() {
  await ensureXlwingsLoaded();

  // Use your actual Codespace URL
  const url = "https://redesigned-space-chainsaw-74jjw9qgwr9frp6r-8000.app.github.dev/interpolation";
  const url2 = "https://redesigned-space-chainsaw-74jjw9qgwr9frp6r-8000.app.github.dev/average";
  const url3 = "https://redesigned-space-chainsaw-74jjw9qgwr9frp6r-8000.app.github.dev/std";
  const url4 = "https://redesigned-space-chainsaw-74jjw9qgwr9frp6r-8000.app.github.dev/Variance";
  const url5 = "https://redesigned-space-chainsaw-74jjw9qgwr9frp6r-8000.app.github.dev/Correlation";
  const url6 = "https://redesigned-space-chainsaw-74jjw9qgwr9frp6r-8000.app.github.dev/Covariance";
  const url7 = "https://redesigned-space-chainsaw-74jjw9qgwr9frp6r-8000.app.github.dev/Hypothesis";

  console.log("Attempting to run xlwings.runPython...");

  try {
    // 1. Change .run to .runPython
    // 2. Remove the Excel.run wrapper
    // 3. Use 'auth' instead of 'headers' (this sends the Authorization header automatically)
    await xlwings.runPython(url, { auth: "DEVELOPMENT" });
    await xlwings.runPython(url2, { auth: "DEVELOPMENT" });
    await xlwings.runPython(url3, { auth: "DEVELOPMENT" });
    await xlwings.runPython(url4, { auth: "DEVELOPMENT" });
    await xlwings.runPython(url5, { auth: "DEVELOPMENT" });
    await xlwings.runPython(url6, { auth: "DEVELOPMENT" });
    await xlwings.runPython(url7, { auth: "DEVELOPMENT" });
    
    console.log("Success! Python function executed.");
  } catch (error) {
    console.error("xlwings runPython error:", error);
  }
}

/**
 * Helper to force-load the xlwings library if the Libraries tab fails
 */
async function ensureXlwingsLoaded() {
  if (typeof xlwings !== 'undefined') return;

  console.log("xlwings not found. Attempting manual injection...");
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = "https://cdn.jsdelivr.net/npm/xlwings@0.30.14/xlwings.min.js";
    script.onload = () => {
      console.log("xlwings loaded successfully via injection.");
      resolve(true);
    };
    script.onerror = () => reject(new Error("Could not load xlwings from CDN. Check your internet or firewall."));
    document.head.appendChild(script);
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error("Error executing script:", error);
  }
}
```

### Script Lab: HTML
Paste the following code into the **HTML** tab:

```html
<button id="run" class="ms-Button">
    <span class="ms-Button-label">Run Forecast</span>
</button>
```

### Script Lab: CSS
Paste the following code into the **CSS** tab:

```css
body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    font-size: 14px;
    line-height: 1.5;
    padding: 10px;
}

section {
    margin-bottom: 20px;
}

h3 {
    margin-top: 0;
    margin-bottom: 10px;
    font-size: 16px;
}

p {
    margin: 0 0 10px 0;
}

button {
    background-color: #f0f0f0;
    color: #333333;
    border: 1px solid #8a8a8a;
    padding: 8px 16px;
    font-size: 14px;
    cursor: pointer;
    border-radius: 2px;
    margin-left: 20px;
    margin-bottom: 5px;
    min-width: 80px;
    display: block;
}

button:hover {
    background-color: #e0e0e0;
}

button:active {
    background-color: #d0d0d0;
}

input {
    padding: 8px;
    margin: 5px 0;
    border: 1px solid #ccc;
    border-radius: 2px;
    font-size: 14px;
}

.header {
    text-align: center;
    background-color: #f3f2f1;
    padding: 10px;
}
```
