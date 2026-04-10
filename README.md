# CurveMaster — Relative Grading System

CurveMaster is a web application that helps professors and teachers grade students on a relative curve. Upload a spreadsheet of student marks, configure weightages, and get automatic grade assignments with adjustable cutoffs, visual histograms, and Excel export.

---

## How to Run This Application on Your Computer

Follow these steps from scratch — even if you've never done this before.

---

### Step 1: Install Node.js

Node.js is the tool that lets you run this application. It also comes with **npm** (Node Package Manager) which is used to install libraries.

1. Open your browser and go to: **https://nodejs.org**
2. You'll see two download buttons — click the one that says **"LTS"** (Long Term Support). This is the stable, recommended version.
3. Once downloaded, open the installer file and follow the on-screen instructions. Just keep clicking **Next / Continue** and accept the defaults.
4. After installation is done, **restart your computer** (recommended, especially on Windows).

#### ✅ Verify it worked

Open a terminal (or Command Prompt on Windows) and type:

```bash
node --version
```

You should see something like `v20.x.x` or `v22.x.x`. If you see a version number, you're good!

Also check npm:

```bash
npm --version
```

You should see something like `10.x.x`. If both commands show version numbers, Node.js is installed correctly.

> **How to open a terminal:**
> - **Windows**: Press `Win + R`, type `cmd`, and press Enter. Or search for "Command Prompt" in the Start menu.
> - **Mac**: Press `Cmd + Space`, type `Terminal`, and press Enter.
> - **Linux**: Press `Ctrl + Alt + T` or find Terminal in your applications.

---

### Step 2: Download This Project

You have two options:

#### Option A: Download as ZIP (easiest)

1. Go to the GitHub page of this project.
2. Click the green **"Code"** button.
3. Click **"Download ZIP"**.
4. Extract (unzip) the downloaded file to a folder on your computer. Remember where you put it.

#### Option B: Clone with Git (if you have Git installed)

Open your terminal and run:

```bash
git clone https://github.com/aniruddhbshreyash-code/Curvemaster.git
```

This will create a folder called `Curvemaster` with all the project files.

---

### Step 3: Open the Project in Your Terminal

You need to navigate your terminal into the project folder.

If you downloaded the ZIP and extracted it to your Desktop, for example:

**On Windows:**
```bash
cd C:\Users\YourName\Desktop\Curvemaster
```

**On Mac/Linux:**
```bash
cd ~/Desktop/Curvemaster
```

> **Tip:** You can also drag and drop the folder into the terminal on Mac to auto-fill the path.

Make sure you're in the right folder by running:

```bash
ls
```

You should see files like `package.json`, `app/`, `README.md`, etc.

---

### Step 4: Install the Project Dependencies

This project uses several libraries (like React, Recharts for charts, XLSX for Excel export, etc.). You need to download them before the app can run.

In your terminal (make sure you're inside the project folder), run:

```bash
npm install
```

This will take a minute or two. You'll see a progress bar and it will download everything the app needs into a folder called `node_modules`.

> **Don't worry** if you see some warnings — that's normal. As long as you don't see red `ERR!` messages, you're fine.

---

### Step 5: Start the Application

Now you're ready to run the app! In the same terminal, run:

```bash
npm run dev
```

You'll see something like this:

```
▲ Next.js 16.x.x
- Local: http://localhost:3000
```

**The app is now running on your computer!**

---

### Step 6: Open the App in Your Browser

Open any web browser (Chrome, Firefox, Edge, Safari — any is fine) and go to:

```
http://localhost:3000
```

You should see the CurveMaster interface with an upload area.

---

### Step 7: Using the App

1. **Upload** — Drop or browse for your student marks spreadsheet (CSV or Excel file).
2. **Configure** — Select the name column, choose which mark columns to include, set weightages and max marks. Optionally enable team projects.
3. **Dashboard** — View the grade distribution histogram, adjust grade cutoffs with sliders, and export the results to Excel.

---

### How to Stop the App

When you're done, go back to the terminal where the app is running and press:

```
Ctrl + C
```

This stops the development server.

---

### How to Start It Again Next Time

You don't need to install anything again. Just:

1. Open your terminal
2. Navigate to the project folder (`cd path/to/Curvemaster`)
3. Run `npm run dev`
4. Open `http://localhost:3000` in your browser

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `node: command not found` | Node.js is not installed or your terminal needs to be restarted. Re-do Step 1. |
| `npm: command not found` | Same as above — npm comes with Node.js. |
| `npm install` shows errors | Delete the `node_modules` folder and the `package-lock.json` file, then run `npm install` again. |
| Page shows blank at localhost:3000 | Make sure `npm run dev` is still running in your terminal. Check for error messages there. |
| Port 3000 is already in use | Another app is using that port. Either close it, or run `npm run dev -- -p 3001` to use port 3001 instead. |

---

## Tech Stack

- **Next.js 16** — React framework
- **React 19** — UI library
- **Recharts** — Charts and histograms
- **XLSX (SheetJS)** — Excel file reading and export
- **Lucide React** — Icons
- **Tailwind CSS 4** — Styling
- **TypeScript** — Type-safe JavaScript
