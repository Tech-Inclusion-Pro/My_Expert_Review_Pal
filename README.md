<p align="center">
  <img src="MERP.png" alt="My Expert Review Pal Logo" width="180" />
</p>

<h1 align="center">My Expert Review Pal (MERP)</h1>

<p align="center">
  <strong>Train your own personal AI review assistant for academic papers, dissertations, and scholarly documents.</strong>
</p>

<p align="center">
  <a href="#-installation">Installation</a> &bull;
  <a href="#-features">Features</a> &bull;
  <a href="#-getting-started">Getting Started</a> &bull;
  <a href="#-how-to-use">How to Use</a> &bull;
  <a href="#-accessibility">Accessibility</a> &bull;
  <a href="#-api-setup">API Setup</a> &bull;
  <a href="#-faq">FAQ</a>
</p>

---

## What is MERP?

**My Expert Review Pal** is a desktop application that helps academics, students, and researchers review scholarly documents with AI-powered analysis. Configure custom review areas, upload your document, and get detailed feedback across writing quality, logical flow, formatting compliance, accessibility, content depth, tables & figures, and mathematical accuracy.

MERP is designed to be **your personal review assistant** that you can train with your own prompts, reference materials, and knowledge examples — so it reviews documents the way *you* would.

---

## Installation

### Download Pre-Built App

Go to the [**Releases**](https://github.com/Tech-Inclusion-Pro/My_Expert_Review_Pal/releases) page and download the version for your operating system:

| Platform | Download | Notes |
|----------|----------|-------|
| **macOS (Apple Silicon)** | `My.Expert.Review.Pal-mac-arm64.zip` | For M1/M2/M3/M4 Macs |
| **macOS (Intel)** | `My.Expert.Review.Pal-mac-x64.zip` | For older Intel-based Macs |
| **Windows** | `My.Expert.Review.Pal-win-x64.zip` | For 64-bit Windows 10/11 |
| **Linux** | `My.Expert.Review.Pal-linux-x64.zip` | For 64-bit Linux (Ubuntu, Fedora, etc.) |

#### macOS Installation

1. Download the `.zip` file for your Mac type (Apple Silicon or Intel)
2. Double-click the zip to extract the `.app` file
3. Drag **My Expert Review Pal.app** into your **Applications** folder
4. On first launch, you may see a security warning:
   - Go to **System Settings > Privacy & Security**
   - Scroll down and click **"Open Anyway"** next to the MERP warning
   - Alternatively, right-click the app and choose **Open** the first time

#### Windows Installation

1. Download `My.Expert.Review.Pal-win-x64.zip`
2. Extract the zip to a folder of your choice (e.g., `C:\Program Files\MERP\`)
3. Double-click **My Expert Review Pal.exe** to launch
4. If Windows Defender SmartScreen appears, click **"More info"** then **"Run anyway"**

#### Linux Installation

1. Download `My.Expert.Review.Pal-linux-x64.zip`
2. Extract the zip:
   ```bash
   unzip My.Expert.Review.Pal-linux-x64.zip
   ```
3. Make the executable runnable:
   ```bash
   chmod +x "My Expert Review Pal"
   ```
4. Launch:
   ```bash
   ./"My Expert Review Pal"
   ```

### Build from Source

If you prefer to build from source:

```bash
# Clone the repository
git clone https://github.com/Tech-Inclusion-Pro/My_Expert_Review_Pal.git
cd My_Expert_Review_Pal

# Install dependencies
npm install

# Run in development mode (opens in browser)
npm run dev

# Build the Electron app for your platform
npm run dist
```

**Requirements:** Node.js 18+ and npm 9+

---

## Features

### 7 Specialized Review Areas

| Area | What It Reviews |
|------|----------------|
| **General Writing** | Grammar, syntax, sentence clarity, academic tone, redundancy, word choice |
| **Flow** | Paragraph transitions, reasoning gaps, argument arc, signposting |
| **Formatting** | Citation style (APA 7, Chicago, MLA, etc.), heading hierarchy, acronyms |
| **Accessibility** | WCAG 2.1 compliance, heading structure, alt text, plain language, inclusive language |
| **Content Area** | Evidence support, citation currency, depth of coverage, counterarguments |
| **Tables & Figures** | Captions, color independence, in-text references, accessibility ratings |
| **Mathematical** | Arithmetic accuracy, statistical reporting, formula notation, consistency |

### Document Support

- **Upload** `.txt`, `.md`, `.csv`, `.pdf`, and `.docx` files directly
- **Paste** text from any source
- Automatic text extraction from PDFs and Word documents

### Configurable AI Prompts

- Edit the prompt for each review area to match your specific needs
- Add **branching logic** — conditional instructions that activate based on document content
- Upload **reference files** (rubrics, style guides, criteria) for context
- Add **knowledge examples** — mark them as "good" or "bad" with notes to train the AI
- Set **priority levels** (Critical / High / Medium / Low) per area and per branch
- Choose **location marking** options (page numbers, paragraph numbers, first 5 words, etc.)

### Export & Reports

- **Save reports** to revisit later
- **Export as HTML** — standalone, printable web page
- **Download as DOCX** — fully accessible Word document with proper headings, bold/italic formatting, and WCAG structure
- **Download Checklist** — abbreviated, priority-sorted checklist DOCX with checkboxes for items to address
- **Generate Combined Prompt** — export your entire configuration as a single prompt for use with other AI tools
- **Download as Markdown** — export configuration as a structured `.md` file

### Pattern Analysis

- Analyze trends **across multiple reviews** over time
- Filter by project name or tags
- Identify common error types, recurring themes, and improvement patterns

### Multi-Provider API Support

| Provider | Description |
|----------|-------------|
| **Anthropic Cloud** | Claude API (default) — `claude-sonnet-4-20250514` |
| **Ollama (Local)** | Run models locally — llama3, mistral, codellama, etc. |
| **Custom Endpoint** | Any OpenAI-compatible API |

### Built-in Accessibility

- **Dark mode** / Light mode / System theme
- **OpenDyslexic font** toggle for dyslexia-friendly reading
- **Bionic Reading** mode — bolds the first half of each word for faster reading
- **Color blindness filters** — Protanopia, Deuteranopia, Tritanopia, Monochrome
- **Adjustable font size** (75%–175%)
- **Large cursor** and **cursor trail** options
- **Draggable accessibility widget** — position it wherever you need
- Full **keyboard navigation** and **screen reader** support
- **Skip to main content** link

---

## Getting Started

### 1. Set Up Your API Key

Before you can run reviews, you need an AI provider:

**Option A: Anthropic Cloud (Recommended)**
1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Create an account and generate an API key
3. In MERP, go to **API Settings** and paste your key

**Option B: Ollama (Free, Local)**
1. Install [Ollama](https://ollama.ai) on your machine
2. Pull a model: `ollama pull llama3`
3. In MERP, go to **API Settings**, select **Ollama (Local)**
4. Set endpoint to `http://localhost:11434` and enter your model name

**Option C: Custom Endpoint**
1. In MERP, go to **API Settings**, select **Custom Endpoint**
2. Enter your endpoint URL, API key, and model name

### 2. Create an Account

When you first open MERP, you will see a login screen:
1. Click **Register** to create an account
2. Choose a username and password
3. Optionally set a 4–6 digit **passcode** for quick login
4. Check **Remember me** to stay logged in between sessions

### 3. Configure Your Review Areas

Go to **Configure Areas** to customize how MERP reviews documents:
1. Select an area from the left panel
2. **Enable/disable** areas you need
3. **Edit the prompt** to match your specific review needs
4. Add **reference files** (rubrics, style guides)
5. Add **knowledge examples** with good/bad ratings
6. Set up **branching logic** for conditional instructions
7. Choose a **priority level** for each area
8. Click **Save All Changes**

---

## How to Use

### Running a Review

1. Click **New Review** in the sidebar
2. **Step 1 — Upload**: Drop a file or paste text
3. **Step 2 — Configure**:
   - Name your project
   - Select a format style (APA 7, Chicago, MLA, etc.)
   - Add tags for organization
   - Choose which review areas to include
   - Add any additional context
4. **Step 3 — Results**: Click **Run Review** and wait for the AI to analyze each area
5. Review the results — each area shows detailed feedback and flags for human review

### Saving and Exporting

After a review completes:
- **Save Report** — stores the report in MERP for later access
- **Export HTML** — downloads a standalone HTML file
- **Download DOCX** — downloads a fully accessible Word document
- **Download Checklist** — downloads a priority-sorted action item checklist

### Using the Dashboard

The **Dashboard** shows all your projects at a glance:
- See total projects, completed reviews, and saved reports
- Search by name, tag, or notes
- Open any project to continue working or re-export results
- Delete projects you no longer need

### Pattern Analysis

In the **Reports** tab:
1. Switch to the **Pattern Analysis** sub-tab
2. Filter by person or tag (optional)
3. Enter specific themes or words to look for
4. Select analysis types (common errors, recurring themes, improvement over time, etc.)
5. Click **Run Analysis** to get insights across all your reviews

---

## Accessibility

MERP was built with accessibility as a core principle. The floating accessibility widget (bottom-right corner) provides:

| Feature | Description |
|---------|-------------|
| **Theme** | Light, Dark, or System-detected theme |
| **Color Vision** | Filters for Protanopia, Deuteranopia, Tritanopia, and Monochrome |
| **Font Size** | Slider from 75% to 175% |
| **Dyslexia Font** | Toggle OpenDyslexic font across the entire app |
| **Bionic Reading** | Bold the first half of each word for faster reading |
| **Cursor Options** | Default, Large Cursor, or Cursor with Trail |

All exported DOCX documents use proper heading structure, semantic formatting, and WCAG-compliant organization.

---

## API Setup

### Anthropic (Default)

MERP uses the Anthropic API by default with the model `claude-sonnet-4-20250514`.

1. Visit [console.anthropic.com](https://console.anthropic.com)
2. Sign up or log in
3. Go to **API Keys** and create a new key
4. Copy the key (starts with `sk-ant-...`)
5. In MERP: **API Settings** > paste the key > click **Save**
6. Click **Test Connection** to verify

**Cost**: Anthropic charges per token. A typical full review of a 5,000-word document across all 7 areas costs approximately $0.10–$0.30.

### Ollama (Free, Local)

Run AI models entirely on your own machine — no API costs, no data leaves your computer.

1. Install Ollama from [ollama.ai](https://ollama.ai)
2. Open a terminal and pull a model:
   ```bash
   ollama pull llama3        # General purpose, good quality
   ollama pull mistral       # Fast, good for shorter docs
   ollama pull codellama     # Better for technical content
   ```
3. In MERP: **API Settings** > select **Ollama (Local)**
4. Endpoint: `http://localhost:11434`
5. Model: enter the model name you pulled (e.g., `llama3`)

### Custom Endpoint

Use any OpenAI-compatible API (OpenAI, Together.ai, Groq, local vLLM, etc.):

1. In MERP: **API Settings** > select **Custom Endpoint**
2. Enter your endpoint URL (e.g., `https://api.openai.com`)
3. Enter your API key (if required)
4. Enter the model name (e.g., `gpt-4o`)

---

## Project Structure

```
My_Expert_Review_Pal/
├── src/
│   ├── App.jsx              # Main application (all components)
│   ├── main.jsx             # React entry point
│   └── assets/
│       ├── MERP.png         # App logo
│       └── fonts/           # OpenDyslexic font files
├── public/
│   └── MERP.png             # Favicon source
├── build/
│   └── icon.icns            # macOS app icon
├── index.html               # HTML entry with dark mode & accessibility CSS
├── electron-main.cjs        # Electron main process
├── vite.config.js           # Vite configuration
├── package.json             # Dependencies and build config
└── README.md
```

---

## FAQ

### Do I need an API key?

Yes — MERP needs an AI provider to analyze documents. You can use:
- **Anthropic** (paid, cloud-based, highest quality)
- **Ollama** (free, runs on your computer)
- **Any OpenAI-compatible API** (varies)

### Is my data sent to the cloud?

Only your document text is sent to the AI provider you configure. If you use Ollama, everything stays on your machine. All project data, settings, and reports are stored locally in your browser's localStorage.

### Can I use this for non-academic documents?

Absolutely. While the default prompts are tuned for academic writing, you can customize every prompt in **Configure Areas** to match any document type — business reports, creative writing, technical documentation, etc.

### How do I back up my data?

MERP stores all data in localStorage within the Electron app. To back up:
1. Open DevTools (Cmd+Option+I on Mac, Ctrl+Shift+I on Windows)
2. Go to **Application** > **Local Storage**
3. Right-click and copy the data

### Can multiple people use the app on one computer?

Yes — MERP has a built-in login system. Each user creates their own account with a username and password. User sessions are isolated.

### What format styles are supported?

APA 7, Chicago, MLA, Harvard, Vancouver, IEEE, AMA, and None/Other. The format style is automatically injected into the Formatting review area prompt.

---

## Tech Stack

- **Frontend**: React 19 + Vite 8
- **Desktop**: Electron 41
- **PDF Extraction**: pdfjs-dist
- **DOCX Reading**: mammoth
- **DOCX Generation**: docx + file-saver
- **Fonts**: OpenDyslexic (self-hosted)
- **AI**: Anthropic Claude API / Ollama / OpenAI-compatible

---

## License

This project is developed by [Tech Inclusion Pro](https://github.com/Tech-Inclusion-Pro).

---

<p align="center">
  <img src="MERP.png" alt="MERP" width="60" />
  <br>
  <em>Built with accessibility and inclusion at its core.</em>
</p>
