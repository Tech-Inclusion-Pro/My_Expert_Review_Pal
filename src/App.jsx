import { useState, useEffect, useCallback, useRef } from 'react';
import logoImg from './assets/MERP.png';
import fontRegularUrl from './assets/fonts/OpenDyslexic-Regular.otf';
import fontBoldUrl from './assets/fonts/OpenDyslexic-Bold.otf';
import fontItalicUrl from './assets/fonts/OpenDyslexic-Italic.otf';
import * as pdfjsLib from 'pdfjs-dist';
import * as mammoth from 'mammoth';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle } from 'docx';
import { saveAs } from 'file-saver';

pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

// ─── SVG LINE ICONS ─────────────────────────────────────────────────────────────
const sz = { width: 18, height: 18, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 2, strokeLinecap: 'round', strokeLinejoin: 'round' };
const Icon = ({ children, size = 18, ...p }) => <svg {...sz} width={size} height={size} {...p}>{children}</svg>;

const Icons = {
  dashboard: (p) => <Icon {...p}><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></Icon>,
  newReview: (p) => <Icon {...p}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></Icon>,
  configure: (p) => <Icon {...p}><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></Icon>,
  apiKey: (p) => <Icon {...p}><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/></Icon>,
  reports: (p) => <Icon {...p}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></Icon>,
  writing: (p) => <Icon {...p}><path d="M17 3a2.85 2.85 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/></Icon>,
  flow: (p) => <Icon {...p}><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></Icon>,
  format: (p) => <Icon {...p}><line x1="21" y1="10" x2="3" y2="10"/><line x1="21" y1="6" x2="3" y2="6"/><line x1="21" y1="14" x2="3" y2="14"/><line x1="21" y1="18" x2="3" y2="18"/></Icon>,
  access: (p) => <Icon {...p}><circle cx="12" cy="4.5" r="2.5"/><path d="M12 7v5"/><path d="m8 17 4-5 4 5"/><path d="M7 22l2.5-5"/><path d="M17 22l-2.5-5"/></Icon>,
  contentArea: (p) => <Icon {...p}><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/><line x1="8" y1="7" x2="16" y2="7"/><line x1="8" y1="11" x2="14" y2="11"/></Icon>,
  tables: (p) => <Icon {...p}><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/><line x1="15" y1="3" x2="15" y2="21"/></Icon>,
  math: (p) => <Icon {...p}><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/><circle cx="12" cy="12" r="10" strokeWidth="1.5"/></Icon>,
  upload: (p) => <Icon {...p}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></Icon>,
  clip: (p) => <Icon {...p}><path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/></Icon>,
  check: (p) => <Icon {...p}><polyline points="20 6 9 17 4 12"/></Icon>,
  x: (p) => <Icon {...p}><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></Icon>,
  clock: (p) => <Icon {...p}><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></Icon>,
  spinner: (p) => <Icon {...p}><path d="M21 12a9 9 0 1 1-6.219-8.56" strokeDasharray="40 20"/></Icon>,
  a11y: (p) => <Icon size={22} {...p}><circle cx="12" cy="4.5" r="2.5"/><path d="M12 7v5"/><path d="m8 17 4-5 4 5"/><path d="M7 22l2.5-5"/><path d="M17 22l-2.5-5"/></Icon>,
  sun: (p) => <Icon {...p}><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></Icon>,
  moon: (p) => <Icon {...p}><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></Icon>,
  monitor: (p) => <Icon {...p}><rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></Icon>,
  eye: (p) => <Icon {...p}><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></Icon>,
  type: (p) => <Icon {...p}><polyline points="4 7 4 4 20 4 20 7"/><line x1="9" y1="20" x2="15" y2="20"/><line x1="12" y1="4" x2="12" y2="20"/></Icon>,
  cursor: (p) => <Icon {...p}><path d="M5 3l14 8-6 2-2 6z"/><line x1="13" y1="13" x2="19" y2="19"/></Icon>,
};

const AREA_ICON_MAP = { writing: Icons.writing, flow: Icons.flow, format: Icons.format, access: Icons.access, content: Icons.contentArea, tables: Icons.tables, math: Icons.math };
const NAV_ICON_MAP = { '/dash': Icons.dashboard, '/review': Icons.newReview, '/cfg': Icons.configure, '/api': Icons.apiKey, '/reports': Icons.reports };

// ─── ACCESSIBILITY WIDGET ───────────────────────────────────────────────────────
const CB_FILTERS = {
  none: 'none',
  protanopia: 'url(#cb-protanopia)',
  deuteranopia: 'url(#cb-deuteranopia)',
  tritanopia: 'url(#cb-tritanopia)',
  monochrome: 'grayscale(100%)',
};

function AccessibilityWidget() {
  const [open, setOpen] = useState(false);
  const [theme, setTheme] = useState(() => localStorage.getItem('a11y_theme') || 'system');
  const [cbMode, setCbMode] = useState(() => localStorage.getItem('a11y_cb') || 'none');
  const [fontSize, setFontSize] = useState(() => parseInt(localStorage.getItem('a11y_fs') || '100', 10));
  const [dyslexic, setDyslexic] = useState(() => localStorage.getItem('a11y_dyslexic') === 'true');
  const [bionic, setBionic] = useState(() => localStorage.getItem('a11y_bionic') === 'true');
  const [cursorMode, setCursorMode] = useState(() => localStorage.getItem('a11y_cursor') || 'default');
  const [dragPos, setDragPos] = useState(() => {
    const saved = localStorage.getItem('a11y_pos');
    return saved ? JSON.parse(saved) : { bottom: 20, right: 20 };
  });
  const [dragging, setDragging] = useState(false);
  const dragRef = useRef(null);

  // persist
  useEffect(() => { localStorage.setItem('a11y_theme', theme); }, [theme]);
  useEffect(() => { localStorage.setItem('a11y_cb', cbMode); }, [cbMode]);
  useEffect(() => { localStorage.setItem('a11y_fs', String(fontSize)); }, [fontSize]);
  useEffect(() => { localStorage.setItem('a11y_dyslexic', String(dyslexic)); }, [dyslexic]);
  useEffect(() => { localStorage.setItem('a11y_bionic', String(bionic)); }, [bionic]);
  useEffect(() => { localStorage.setItem('a11y_cursor', cursorMode); }, [cursorMode]);
  useEffect(() => { localStorage.setItem('a11y_pos', JSON.stringify(dragPos)); }, [dragPos]);

  // apply theme
  useEffect(() => {
    const root = document.documentElement;
    root.removeAttribute('data-theme');
    if (theme === 'dark') root.setAttribute('data-theme', 'dark');
    else if (theme === 'light') root.setAttribute('data-theme', 'light');
  }, [theme]);

  // apply font size
  useEffect(() => { document.documentElement.style.fontSize = `${fontSize}%`; }, [fontSize]);

  // apply dyslexic font (dynamically inject @font-face with Vite-resolved URLs)
  useEffect(() => {
    if (dyslexic) {
      // Inject @font-face if not already present
      if (!document.getElementById('opendyslexic-faces')) {
        const style = document.createElement('style');
        style.id = 'opendyslexic-faces';
        style.textContent = `
          @font-face { font-family: 'OpenDyslexic'; src: url('${fontRegularUrl}') format('opentype'); font-weight: 400; font-style: normal; font-display: swap; }
          @font-face { font-family: 'OpenDyslexic'; src: url('${fontBoldUrl}') format('opentype'); font-weight: 700; font-style: normal; font-display: swap; }
          @font-face { font-family: 'OpenDyslexic'; src: url('${fontItalicUrl}') format('opentype'); font-weight: 400; font-style: italic; font-display: swap; }
        `;
        document.head.appendChild(style);
      }
      document.documentElement.style.fontFamily = "'OpenDyslexic', Arial, sans-serif";
    } else {
      document.documentElement.style.fontFamily = '';
    }
    window.dispatchEvent(new Event('a11ychange'));
  }, [dyslexic]);

  // apply bionic reading mode
  useEffect(() => {
    if (bionic) {
      document.documentElement.classList.add('bionic-active');
    } else {
      document.documentElement.classList.remove('bionic-active');
    }
    window.dispatchEvent(new Event('a11ychange'));
  }, [bionic]);

  // apply color blindness filter
  useEffect(() => {
    const root = document.getElementById('root');
    if (!root) return;
    if (cbMode === 'monochrome') { root.style.filter = 'grayscale(100%)'; }
    else if (cbMode !== 'none') { root.style.filter = CB_FILTERS[cbMode] || 'none'; }
    else { root.style.filter = 'none'; }
  }, [cbMode]);

  // apply cursor
  useEffect(() => {
    document.body.style.cursor = cursorMode === 'large' ? 'url("data:image/svg+xml,%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 width=%2748%27 height=%2748%27%3E%3Cpath d=%27M8 4l28 16-12 4-4 12z%27 fill=%27%23000%27 stroke=%27%23fff%27 stroke-width=%272%27/%3E%3C/svg%3E") 8 4, auto' : '';
    // cursor trail
    const handleMove = (e) => {
      if (cursorMode !== 'trail') return;
      const dot = document.createElement('div');
      dot.className = 'cursor-trail-dot';
      dot.style.cssText = `position:fixed;left:${e.clientX}px;top:${e.clientY}px;width:8px;height:8px;background:${C.accent};border-radius:50%;pointer-events:none;z-index:99999;opacity:0.7;transition:opacity 0.5s,transform 0.5s;`;
      document.body.appendChild(dot);
      setTimeout(() => { dot.style.opacity = '0'; dot.style.transform = 'scale(0)'; }, 50);
      setTimeout(() => dot.remove(), 550);
    };
    if (cursorMode === 'trail') document.addEventListener('mousemove', handleMove);
    return () => { document.removeEventListener('mousemove', handleMove); document.body.style.cursor = ''; };
  }, [cursorMode]);

  // drag
  const onDragStart = (e) => {
    e.preventDefault();
    setDragging(true);
    const startX = e.clientX, startY = e.clientY;
    const startPos = { ...dragPos };
    const onMove = (ev) => {
      setDragPos({ bottom: startPos.bottom - (ev.clientY - startY), right: startPos.right - (ev.clientX - startX) });
    };
    const onUp = () => { setDragging(false); document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  };

  const panelStyle = {
    position: 'fixed', bottom: dragPos.bottom + 56, right: dragPos.right, zIndex: 10002,
    background: 'var(--a11y-panel-bg, #fff)', border: `2px solid ${C.primary}`, borderRadius: 12,
    padding: '1.25rem', width: 300, maxHeight: '80vh', overflowY: 'auto',
    boxShadow: '0 8px 32px rgba(0,0,0,0.18)', color: 'var(--a11y-panel-text, #1a1a2e)',
  };

  const secStyle = { marginBottom: 14, paddingBottom: 10, borderBottom: `1px solid ${C.border}` };
  const secTitle = { fontWeight: 700, fontSize: '0.85rem', marginBottom: 6, display: 'flex', alignItems: 'center', gap: 6 };
  const optBtn = (active) => ({
    padding: '0.3rem 0.7rem', borderRadius: 5, border: `1.5px solid ${active ? C.primary : C.border}`,
    background: active ? C.primary : 'transparent', color: active ? '#fff' : 'inherit',
    cursor: 'pointer', fontSize: '0.8rem', fontFamily: 'Arial, sans-serif', fontWeight: active ? 600 : 400,
  });

  return (
    <>
      {/* SVG filters for color blindness */}
      <svg style={{ position: 'absolute', width: 0, height: 0 }} aria-hidden="true">
        <defs>
          <filter id="cb-protanopia"><feColorMatrix type="matrix" values="0.567,0.433,0,0,0 0.558,0.442,0,0,0 0,0.242,0.758,0,0 0,0,0,1,0"/></filter>
          <filter id="cb-deuteranopia"><feColorMatrix type="matrix" values="0.625,0.375,0,0,0 0.7,0.3,0,0,0 0,0.3,0.7,0,0 0,0,0,1,0"/></filter>
          <filter id="cb-tritanopia"><feColorMatrix type="matrix" values="0.95,0.05,0,0,0 0,0.433,0.567,0,0 0,0.475,0.525,0,0 0,0,0,1,0"/></filter>
        </defs>
      </svg>

      {/* Floating button */}
      <button ref={dragRef} onMouseDown={onDragStart}
        onClick={() => { if (!dragging) setOpen(!open); }}
        aria-label="Accessibility settings" aria-expanded={open}
        style={{
          position: 'fixed', bottom: dragPos.bottom, right: dragPos.right, zIndex: 10002,
          width: 48, height: 48, borderRadius: '50%', border: `2px solid #fff`,
          background: `linear-gradient(135deg, ${C.primary}, ${C.accent})`, color: '#fff',
          cursor: 'grab', display: 'flex', alignItems: 'center', justifyContent: 'center',
          boxShadow: '0 4px 16px rgba(0,0,0,0.25)', padding: 0,
        }}>
        <Icons.a11y style={{ color: '#fff' }} />
      </button>

      {/* Panel */}
      {open && (
        <div style={panelStyle} role="dialog" aria-label="Accessibility settings panel">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
            <h2 style={{ fontSize: '1rem', fontWeight: 700, color: C.primary, margin: 0 }}>Accessibility</h2>
            <button onClick={() => setOpen(false)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'inherit', padding: 2 }} aria-label="Close accessibility panel"><Icons.x /></button>
          </div>

          {/* Theme */}
          <div style={secStyle}>
            <div style={secTitle}><Icons.sun size={16}/> Appearance</div>
            <div style={{ display: 'flex', gap: 6 }}>
              {[['light','Light',Icons.sun],['dark','Dark',Icons.moon],['system','System',Icons.monitor]].map(([v,l,I]) => (
                <button key={v} onClick={() => setTheme(v)} style={optBtn(theme===v)} aria-pressed={theme===v}>
                  <I size={14} style={{ verticalAlign: 'middle', marginRight: 4 }}/>{l}
                </button>
              ))}
            </div>
          </div>

          {/* Color blindness */}
          <div style={secStyle}>
            <div style={secTitle}><Icons.eye size={16}/> Color Vision</div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 5 }}>
              {[['none','Normal'],['protanopia','Protanopia'],['deuteranopia','Deuteranopia'],['tritanopia','Tritanopia'],['monochrome','Monochrome']].map(([v,l]) => (
                <button key={v} onClick={() => setCbMode(v)} style={optBtn(cbMode===v)} aria-pressed={cbMode===v}>{l}</button>
              ))}
            </div>
          </div>

          {/* Font size */}
          <div style={secStyle}>
            <div style={secTitle}><Icons.type size={16}/> Font Size ({fontSize}%)</div>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <button onClick={() => setFontSize(f => Math.max(75, f - 10))} style={optBtn(false)} aria-label="Decrease font size">A-</button>
              <input type="range" min="75" max="175" step="5" value={fontSize} onChange={e => setFontSize(+e.target.value)}
                style={{ flex: 1, accentColor: C.primary }} aria-label="Font size slider"/>
              <button onClick={() => setFontSize(f => Math.min(175, f + 10))} style={optBtn(false)} aria-label="Increase font size">A+</button>
            </div>
          </div>

          {/* Dyslexic font */}
          <div style={secStyle}>
            <div style={secTitle}><Icons.type size={16}/> Dyslexia-Friendly Font</div>
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
              <input type="checkbox" checked={dyslexic} onChange={e => setDyslexic(e.target.checked)} style={{ accentColor: C.primary, width: 18, height: 18 }}/>
              Use OpenDyslexic font
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.85rem', marginTop: 8 }}>
              <input type="checkbox" checked={bionic} onChange={e => setBionic(e.target.checked)} style={{ accentColor: C.primary, width: 18, height: 18 }}/>
              Bionic Reading mode
            </label>
          </div>

          {/* Cursor */}
          <div style={{ marginBottom: 0 }}>
            <div style={secTitle}><Icons.cursor size={16}/> Cursor Options</div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 5 }}>
              {[['default','Default'],['large','Large Cursor'],['trail','Cursor with Trail']].map(([v,l]) => (
                <button key={v} onClick={() => setCursorMode(v)} style={optBtn(cursorMode===v)} aria-pressed={cursorMode===v}>{l}</button>
              ))}
            </div>
          </div>

          {/* Reset */}
          <button onClick={() => { setTheme('system'); setCbMode('none'); setFontSize(100); setDyslexic(false); setBionic(false); setCursorMode('default'); }}
            style={{ ...optBtn(false), marginTop: 14, width: '100%', textAlign: 'center', color: C.error, borderColor: C.error }}>
            Reset All to Defaults
          </button>
        </div>
      )}
    </>
  );
}

// ─── CONSTANTS ──────────────────────────────────────────────────────────────────
const C = {
  primary: '#3a2b95', secondary: '#a23b84', accent: '#6f2fa6',
  bg: '#f4f2fb', card: '#ffffff', border: '#ddd8f5', text: '#1a1a2e',
  success: '#2e7d32', error: '#c62828', warn: '#f9a825', info: '#1565c0',
  flagBg: '#fff8e1', flagBorder: '#ffca28', flagText: '#5d4037',
};

const AREAS = [
  { id: 'writing', label: 'General Writing', description: 'Grammar, syntax, structure, and academic tone' },
  { id: 'flow', label: 'Flow', description: 'Logical progression and transitions between ideas' },
  { id: 'format', label: 'Formatting', description: 'Citation style, headings, and document conventions' },
  { id: 'access', label: 'Accessibility', description: 'WCAG 2.1 compliance and inclusive language' },
  { id: 'content', label: 'Content Area', description: 'Accuracy, depth, and relevance of subject matter' },
  { id: 'tables', label: 'Tables & Figures', description: 'WCAG compliance, captions, and visual clarity' },
  { id: 'math', label: 'Mathematical', description: 'Arithmetic, statistical, and formula accuracy' },
];
const AreaIcon = ({ id, ...p }) => { const I = AREA_ICON_MAP[id]; return I ? <I {...p}/> : null; };

const FMT_STYLES = ['APA 7','Chicago','MLA','Harvard','Vancouver','IEEE','AMA','None/Other'];

const DEFAULT_PROMPTS = {
  writing: `You are an expert academic writing reviewer. Analyze the provided document for the following writing quality dimensions:

1. **Grammar & Syntax**: Identify grammatical errors, subject-verb agreement issues, misplaced modifiers, comma splices, run-on sentences, and sentence fragments.
2. **Sentence Clarity**: Flag overly complex, convoluted, or ambiguous sentences. Suggest clearer alternatives.
3. **Paragraph Coherence**: Evaluate whether each paragraph has a clear topic sentence, supporting details, and a logical conclusion. Flag paragraphs that try to cover too many ideas.
4. **Academic Tone**: Identify informal language, colloquialisms, contractions, or first-person usage that may be inappropriate for the document type. Flag emotional or biased language.
5. **Redundancy**: Flag repeated ideas, phrases, or words that appear unnecessarily. Identify wordiness and suggest concise alternatives.
6. **Active vs. Passive Voice**: Flag excessive passive voice usage. Note where passive voice is appropriate (e.g., methods sections) vs. where active voice would strengthen the writing.
7. **Terminology Consistency**: Identify instances where the same concept is referred to by different terms without explanation. Flag undefined jargon or acronyms.
8. **Word Choice**: Flag vague words (e.g., "things," "stuff," "very," "really"), clichés, and imprecise language. Suggest specific, scholarly alternatives.

For each issue found, report:
- **Location**: Page number or section where the issue occurs
- **Context**: The first 5 words of the relevant sentence
- **Issue**: Clear description of the problem
- **Recommendation**: Specific suggestion for improvement

Celebrate strong sections — note particularly well-crafted paragraphs, effective use of academic voice, or excellent clarity.

End your review with a section titled exactly:
⚑ FLAGS FOR HUMAN REVIEW
List anything that requires the human reviewer's personal judgment, such as disciplinary conventions, stylistic preferences, or context-dependent word choices.`,

  flow: `You are an expert in academic document structure and argumentation flow. Analyze the provided document for logical progression and coherence:

1. **Paragraph-to-Paragraph Transitions**: Evaluate whether each paragraph connects logically to the next. Flag abrupt shifts where a transition sentence is missing or ineffective.
2. **Section Bridges**: Check that each major section flows into the next with appropriate bridging language. Flag sections that feel disconnected.
3. **Reasoning Gaps**: Identify places where the argument jumps to a conclusion without sufficient intermediate reasoning or evidence. Flag logical leaps.
4. **Abrupt Topic Shifts**: Flag any place where the topic changes without adequate preparation or signposting for the reader.
5. **Missing "So What" Statements**: Identify sections or paragraphs that present information without explaining its significance to the overall argument or research question.
6. **Argument Arc**: Evaluate whether the document builds a coherent argument from introduction to conclusion. Does the introduction set up expectations that the body fulfills? Does the conclusion synthesize rather than merely summarize?
7. **Signposting**: Check for adequate use of discourse markers and signposting language (e.g., "This section will...", "Having established X, we now turn to...").
8. **Internal Consistency**: Flag any contradictions between sections or claims made early that are later contradicted or forgotten.

For each issue found, report:
- **Location**: Page number or section where the issue occurs
- **Context**: The first 5 words of the relevant sentence
- **Issue**: Clear description of the flow problem
- **Recommendation**: Specific suggestion for improving coherence

Celebrate strong sections — note particularly effective transitions, well-structured arguments, or masterful signposting.

End your review with a section titled exactly:
⚑ FLAGS FOR HUMAN REVIEW
List anything that requires the human reviewer's personal judgment, such as whether certain organizational choices serve a disciplinary purpose or whether unconventional structure is intentional.`,

  format: `You are an expert academic formatting reviewer. Analyze the provided document for formatting compliance and consistency:

1. **Citation Formatting (In-text)**: Check every in-text citation for correct formatting according to the specified style guide. Flag inconsistencies in author names, years, page numbers, and punctuation.
2. **Citation Formatting (Reference List)**: Check every reference list entry for completeness and correct formatting — author names, dates, titles, journal names, volume/issue numbers, DOIs, URLs, and proper punctuation/italicization.
3. **Heading Hierarchy**: Verify that headings follow a logical hierarchy (e.g., Heading 1 → Heading 2 → Heading 3) without skipping levels. Check that heading formatting matches the style guide.
4. **Acronym Definitions**: Verify that every acronym is defined on first use in the format: "Full Term (ACRONYM)". Flag acronyms used without prior definition.
5. **Number & Percentage Formatting**: Check that numbers and percentages follow style guide conventions (e.g., spelling out numbers under 10 in APA, using % vs. "percent").
6. **Table & Figure Captions**: Verify that every table and figure has a properly formatted caption per the style guide. Check numbering sequence (Table 1, Table 2...).
7. **Running Headers/Footers**: If applicable, check that running headers match the style guide requirements.
8. **Verb Tense Consistency**: Check that appropriate verb tenses are used consistently within sections (e.g., past tense for methods/results, present tense for discussion of established findings).
9. **Page Formatting**: Check margins, line spacing, font requirements, and page numbering as applicable.

For each issue found, report:
- **Location**: Page number or section where the issue occurs
- **Context**: The first 5 words of the relevant sentence or the specific citation/reference
- **Issue**: Clear description of the formatting error
- **Recommendation**: Specific correction with the exact formatting required

Celebrate strong sections — note areas with impeccable formatting compliance or particularly well-organized reference lists.

End your review with a section titled exactly:
⚑ FLAGS FOR HUMAN REVIEW
List anything that requires the human reviewer's personal judgment, such as ambiguous style guide interpretations, newer sources that may have updated formatting rules, or institutional-specific requirements.`,

  access: `You are an expert in document accessibility and WCAG 2.1 compliance. Analyze the provided document for accessibility:

1. **Heading Hierarchy (WCAG 1.3.1 — Info and Relationships)**: Verify headings never skip levels (e.g., jumping from h1 to h3). Each heading level must be properly nested.
2. **Alternative Text for Images (WCAG 1.1.1 — Non-text Content)**: Check that all described images, charts, and diagrams have adequate alternative text descriptions. Flag any visual content referenced without accessible descriptions.
3. **Table Structure (WCAG 1.3.1)**: Verify tables have clear headers, row/column labels, and captions. Flag tables that would be difficult to navigate with a screen reader.
4. **Color Independence (WCAG 1.4.1 — Use of Color)**: Identify any place where information is conveyed by color alone without a text alternative. Check that meaning does not rely solely on visual formatting.
5. **Plain Language & Readability (WCAG 3.1.5 — Reading Level)**: Evaluate overall readability. Flag unnecessarily complex sentences, excessive jargon without definition, and sections that could be simplified without losing meaning.
6. **Link Text Clarity (WCAG 2.4.4 — Link Purpose)**: If the document contains hyperlinks, verify that link text is descriptive (not "click here" or bare URLs).
7. **Logical Reading Order (WCAG 1.3.2 — Meaningful Sequence)**: Evaluate whether the document would make sense when read linearly. Flag content whose meaning depends on visual positioning.
8. **Language Identification (WCAG 3.1.1/3.1.2)**: Flag any sections in a language different from the main document language that are not identified as such.
9. **Inclusive Language**: Flag language that may exclude or marginalize groups. Suggest person-first or identity-first alternatives as appropriate.
10. **Abbreviations and Acronyms (WCAG 3.1.4)**: Ensure abbreviations are expanded on first use for screen reader compatibility.

For each issue found, report:
- **Location**: Page number or section where the issue occurs
- **WCAG Criterion**: The specific WCAG 2.1 criterion (e.g., 1.3.1, 1.4.1)
- **Context**: The first 5 words of the relevant sentence or description of the element
- **Issue**: Clear description of the accessibility barrier
- **Recommendation**: Specific suggestion for remediation

Celebrate strong sections — note areas with excellent accessibility practices, well-structured tables, or particularly clear writing.

End your review with a section titled exactly:
⚑ FLAGS FOR HUMAN REVIEW
List anything that requires the human reviewer's personal judgment, such as whether certain complex terminology is necessary for the audience, discipline-specific formatting that may conflict with accessibility, or subjective readability assessments.`,

  content: `You are an expert academic content reviewer. Analyze the provided document for content quality and scholarly rigor:

1. **Accuracy of Claims**: Evaluate whether factual claims appear accurate based on general scholarly knowledge. Flag claims that seem incorrect, outdated, or misleading.
2. **Evidence Support**: For each major claim or argument, assess whether adequate evidence (citations, data, examples) is provided. Flag unsupported claims or assertions presented as fact without evidence.
3. **Citation Currency**: Flag any cited sources that are more than 10 years old and assess whether more current sources should be used. Note: some seminal works are appropriately old — flag these for human judgment.
4. **Depth of Coverage**: Evaluate whether key topics are covered with sufficient depth. Flag sections that are superficial or that gloss over important complexities.
5. **Coverage Gaps**: Identify topics that seem missing or underrepresented given the document's stated scope or research question. Flag areas where additional discussion is expected.
6. **Research Question Alignment**: Evaluate whether the content directly addresses the stated research question, thesis, or purpose. Flag tangential content that does not clearly serve the main argument.
7. **Source Diversity**: Assess whether the document draws on a diverse range of sources (multiple authors, journals, perspectives) or relies too heavily on a small number of sources.
8. **Counterarguments**: Check whether the document acknowledges alternative viewpoints or counterarguments. Flag one-sided arguments that would benefit from balanced perspective.
9. **Definitions and Context**: Verify that key concepts are adequately defined and contextualized for the target audience.

For each issue found, report:
- **Location**: Page number or section where the issue occurs
- **Context**: The first 5 words of the relevant sentence
- **Issue**: Clear description of the content concern
- **Recommendation**: Specific suggestion for improvement, including types of sources to seek or topics to expand

Celebrate strong sections — note particularly well-argued points, excellent use of evidence, or impressive depth of analysis.

End your review with a section titled exactly:
⚑ FLAGS FOR HUMAN REVIEW
List anything that requires the human reviewer's personal judgment, such as disciplinary debates, whether seminal older sources are appropriate, or content accuracy claims that require subject-matter expertise to verify.`,

  tables: `You are an expert in evaluating tables, figures, and visual data presentations in academic documents. Analyze every table and figure in the provided document:

For EACH table, evaluate:
1. **Title/Caption**: Is there a clear, descriptive title above the table? Does it follow the specified style guide format? Does it stand alone (reader can understand the table without reading the full text)?
2. **Column and Row Headers**: Are all columns and rows clearly labeled? Are headers descriptive and unambiguous?
3. **Notes Section**: Does the table include necessary notes (general notes, specific notes, probability notes) as required by the style guide?
4. **Citation Compliance**: If the table contains data from other sources, is it properly cited? If the table is reproduced, is permission noted?
5. **Color Independence**: Would the table be fully readable in black and white? Flag any information conveyed solely through color coding.
6. **In-text Reference**: Is every table referenced in the text (e.g., "as shown in Table 1")? Flag orphan tables with no text reference.
7. **Unit Labels**: Are all measurements clearly labeled with units? Are units consistent?
8. **Accessibility Rating**: Rate each table as HIGH / MEDIUM / LOW accessibility needs based on complexity, header clarity, and screen-reader navigability.

For EACH figure, evaluate:
1. **Caption**: Is there a clear, descriptive caption below the figure? Does it follow the style guide?
2. **Alt Text Potential**: Could the figure be adequately described in alternative text? Flag overly complex figures that need detailed descriptions.
3. **Color Independence**: Would the figure be fully readable in black and white or grayscale?
4. **In-text Reference**: Is the figure referenced in the text?
5. **Labels and Legends**: Are axes labeled, legends included, and all visual elements explained?
6. **Accessibility Rating**: Rate each figure as HIGH / MEDIUM / LOW accessibility needs.

For each issue found, report:
- **Location**: The specific table or figure number and page
- **Context**: Description of the specific element
- **Issue**: Clear description of the problem
- **Recommendation**: Specific suggestion for improvement

Celebrate strong sections — note particularly well-designed tables or figures that serve as good examples.

End your review with a section titled exactly:
⚑ FLAGS FOR HUMAN REVIEW
List anything that requires the human reviewer's personal judgment, such as whether certain visual design choices serve a specific purpose, discipline-specific table conventions, or subjective clarity assessments.`,

  math: `You are an expert in statistical reporting and mathematical accuracy in academic documents. Analyze the provided document for mathematical and statistical correctness:

1. **Arithmetic Accuracy**: Verify any calculations, sums, percentages, or derived numbers mentioned in the text. Flag any apparent calculation errors.
2. **Formula Notation**: Check that all formulas and equations use correct mathematical notation. Flag inconsistent or non-standard notation.
3. **Complete Statistical Reporting**: For each statistical test reported, verify that it includes ALL required elements:
   - Test statistic (e.g., t, F, χ², r, z)
   - Degrees of freedom (df)
   - p-value (exact when possible)
   - Effect size (Cohen's d, η², r², etc.)
   - Confidence intervals (when appropriate)
   Flag any tests with incomplete reporting.
4. **Appropriate Test Selection**: Based on the described data and research design, evaluate whether the correct statistical test was used. Flag potentially inappropriate test choices.
5. **Text-Table-Figure Consistency**: Cross-check all numerical values reported in the text against values in tables and figures. Flag any discrepancies.
6. **Rounding Consistency**: Check that numbers are rounded consistently throughout (e.g., always 2 decimal places for p-values, consistent decimal places for means). Flag inconsistencies.
7. **Descriptive Statistics Completeness**: Verify that descriptive statistics include means, standard deviations (or standard errors), and sample sizes as appropriate. Flag missing descriptive statistics.
8. **Statistical Notation**: Verify correct use of statistical notation:
   - Italicized test statistics (p, t, F, N, M, SD)
   - Correct symbols (≤, ≥, =, <, >)
   - Proper formatting per style guide
9. **Sample Size Reporting**: Verify that sample sizes are clearly reported for all analyses and subgroups.
10. **Assumptions Reporting**: Check whether statistical test assumptions are addressed (normality, homogeneity of variance, independence, etc.).

For each issue found, report:
- **Location**: Page number or section, and the specific statistic/formula
- **Context**: The first 5 words of the relevant sentence
- **Issue**: Clear description of the mathematical or statistical error
- **Recommendation**: Specific correction or the missing information needed

Celebrate strong sections — note particularly thorough statistical reporting, correct and complete analyses, or excellent data presentation.

End your review with a section titled exactly:
⚑ FLAGS FOR HUMAN REVIEW
List anything that requires the human reviewer's personal judgment, such as whether certain statistical approaches are appropriate for the specific research design, discipline-specific reporting conventions, or results that seem plausible but cannot be verified without raw data.`,
};

const AREA_TIPS = {
  writing: 'Describe what writing quality matters most for this document type. Specify tone expectations. Add rubric criteria if you have one.',
  flow: 'Mention the expected argument structure (funnel intro, IMRaD, etc.). Specify document type as flow expectations differ.',
  format: 'The format style selected on the Review page is injected here automatically. Add journal-specific or institutional style variations.',
  access: "Specify the intended audience's accessibility needs. Add institutional accessibility requirements or WCAG checklists.",
  content: 'Upload reference materials (rubrics, course objectives, discipline-specific criteria) so the AI can evaluate content alignment.',
  tables: "Add your style guide's specific table/figure formatting rules. Upload examples of well-formatted tables as reference.",
  math: 'Specify expected statistical software output format (SPSS, R, Stata). Add field-specific reporting standards.',
};

// ─── UTILS ──────────────────────────────────────────────────────────────────────
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

function bionicize(text) {
  if (!text) return text;
  return text.replace(/\b([A-Za-z]+)\b/g, (word) => {
    const mid = Math.ceil(word.length / 2);
    return `<strong>${word.slice(0, mid)}</strong>${word.slice(mid)}`;
  });
}

function BionicText({ children }) {
  const [active, setActive] = useState(() => localStorage.getItem('a11y_bionic') === 'true');
  useEffect(() => {
    const handler = () => setActive(localStorage.getItem('a11y_bionic') === 'true');
    window.addEventListener('a11ychange', handler);
    return () => window.removeEventListener('a11ychange', handler);
  }, []);
  if (!active || typeof children !== 'string') return children;
  return <span dangerouslySetInnerHTML={{ __html: bionicize(children) }} />;
}

// ─── AUTH UTILS ─────────────────────────────────────────────────────────────────
async function hashStr(str) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
}

function getUsers() { return lsGet('merp_users', []); }
function setUsers(users) { lsSet('merp_users', users); }

function getCurrentUser() {
  const s = sessionStorage.getItem('merp_current_user');
  if (s) return JSON.parse(s);
  const l = localStorage.getItem('merp_current_user');
  if (l) return JSON.parse(l);
  return null;
}

function setCurrentUser(user, remember) {
  const data = JSON.stringify(user);
  sessionStorage.setItem('merp_current_user', data);
  if (remember) localStorage.setItem('merp_current_user', data);
  else localStorage.removeItem('merp_current_user');
}

function logoutUser() {
  sessionStorage.removeItem('merp_current_user');
  localStorage.removeItem('merp_current_user');
}

function lsGet(key, fallback = null) {
  try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; }
  catch { return fallback; }
}
function lsSet(key, val) { localStorage.setItem(key, JSON.stringify(val)); }

function getProjects() { return lsGet('prj_list', []); }
function setProjects(list) { lsSet('prj_list', list); }
function getProject(id) { return lsGet(`prj_${id}`); }
function saveProjectData(prj) { lsSet(`prj_${prj.id}`, prj); }
function deleteProject(id) {
  localStorage.removeItem(`prj_${id}`);
  setProjects(getProjects().filter(p => p !== id));
}

function getReports() { return lsGet('rpt_list', []); }
function setReports(list) { lsSet('rpt_list', list); }
function getReport(id) { return lsGet(`rpt_${id}`); }
function saveReportData(rpt) { lsSet(`rpt_${rpt.id}`, rpt); }
function deleteReport(id) {
  localStorage.removeItem(`rpt_${id}`);
  setReports(getReports().filter(r => r !== id));
}

function getAnalyses() { return lsGet('analysis_list', []); }
function setAnalyses(list) { lsSet('analysis_list', list); }
function getAnalysis(id) { return lsGet(`analysis_${id}`); }
function saveAnalysisData(a) { lsSet(`analysis_${a.id}`, a); }
function deleteAnalysis(id) {
  localStorage.removeItem(`analysis_${id}`);
  setAnalyses(getAnalyses().filter(a => a !== id));
}

function getAreaCfg() {
  const stored = lsGet('area_cfg', null);
  if (stored) return stored;
  return Object.fromEntries(AREAS.map(a => [a.id, {
    enabled: true, prompt: DEFAULT_PROMPTS[a.id], files: [], branches: [], knowledgeExamples: [], locationMarking: [], priorityLevel: 'Medium',
  }]));
}
function setAreaCfg(cfg) { lsSet('area_cfg', cfg); }

function getApiKey() { return localStorage.getItem('api_key') || ''; }
function setApiKeyStorage(k) { localStorage.setItem('api_key', k); }

function getApiProvider() {
  return lsGet('api_provider', { type: 'anthropic', endpoint: '', model: '', apiKey: '' });
}
function setApiProvider(p) { lsSet('api_provider', p); }

function splitFlags(text) {
  const marker = '⚑ FLAGS FOR HUMAN REVIEW';
  const idx = text.indexOf(marker);
  if (idx === -1) return { body: text, flags: '' };
  return { body: text.slice(0, idx).trim(), flags: text.slice(idx + marker.length).trim() };
}

function truncateText(s, n) { return s.length > n ? s.slice(0, n) + '\n\n[Document truncated at ' + n + ' characters]' : s; }

function wordCount(s) { return s.trim() ? s.trim().split(/\s+/).length : 0; }

function flattenBranches(branches, depth = 0) {
  let result = '';
  for (const b of branches) {
    const indent = '  '.repeat(depth);
    if (b.condition && b.instructions) {
      const pri = b.priority ? ` [Priority: ${b.priority}]` : '';
      result += `\n${indent}Conditional Branch (Level ${depth + 1})${pri}: When/If ${b.condition} → Apply these additional instructions: ${b.instructions}`;
    }
    if (b.children && b.children.length > 0) {
      result += flattenBranches(b.children, depth + 1);
    }
  }
  return result;
}

function buildSystemPrompt(areaId, areaCfg, fmtStyle, extraCtx) {
  const cfg = areaCfg[areaId] || {};
  let prompt = cfg.prompt || DEFAULT_PROMPTS[areaId];
  if (cfg.priorityLevel && cfg.priorityLevel !== 'Medium') {
    prompt += `\n\nThis review area has been assigned a ${cfg.priorityLevel} priority level. ${cfg.priorityLevel === 'Critical' ? 'Treat every issue in this area as high-stakes. Be exhaustive and flag even minor concerns.' : cfg.priorityLevel === 'High' ? 'Pay extra attention to issues in this area and be thorough.' : 'Focus on major issues only; minor issues can be noted briefly.'}`;
  }
  if (areaId === 'format' && fmtStyle && fmtStyle !== 'None/Other') {
    prompt += `\n\nThis document uses ${fmtStyle}. Apply ${fmtStyle}-specific rules throughout your review.`;
  }
  if (cfg.branches && cfg.branches.length > 0) {
    prompt += '\n\n--- CONDITIONAL BRANCHES ---' + flattenBranches(cfg.branches);
  }
  if (cfg.files && cfg.files.length > 0) {
    for (const f of cfg.files) {
      prompt += `\n\n--- REFERENCE FILE: ${f.name} ---\n${truncateText(f.content, 3000)}`;
    }
  }
  if (cfg.knowledgeExamples && cfg.knowledgeExamples.length > 0) {
    for (const ex of cfg.knowledgeExamples) {
      const label = ex.sentiment === 'liked' ? 'GOOD' : 'BAD';
      prompt += `\n\n--- KNOWLEDGE EXAMPLE (${label}): ${ex.name} ---\n${truncateText(ex.content, 2000)}`;
      if (ex.notes) prompt += `\nReviewer notes: ${ex.notes}`;
    }
  }
  if (cfg.locationMarking && cfg.locationMarking.length > 0) {
    prompt += `\n\n--- LOCATION MARKING INSTRUCTIONS ---\nWhen reporting issues, identify each location using the following method(s): ${cfg.locationMarking.join(', ')}. Use ALL selected methods for every issue you report.`;
  }
  if (extraCtx) {
    prompt += `\n\n--- ADDITIONAL CONTEXT ---\n${extraCtx}`;
  }
  return prompt;
}

async function callApi(apiKey, systemPrompt, userMessage) {
  const provider = getApiProvider();

  if (provider.type === 'ollama') {
    const endpoint = (provider.endpoint || 'http://localhost:11434').replace(/\/$/, '');
    const res = await fetch(`${endpoint}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: provider.model || 'llama3',
        stream: false,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userMessage },
        ],
      }),
    });
    if (!res.ok) {
      const err = await res.text().catch(() => '');
      throw new Error(err || `Ollama error: ${res.status}`);
    }
    const data = await res.json();
    return data.message?.content || '';
  }

  if (provider.type === 'custom') {
    const endpoint = (provider.endpoint || '').replace(/\/$/, '');
    const headers = { 'Content-Type': 'application/json' };
    if (provider.apiKey) headers['Authorization'] = `Bearer ${provider.apiKey}`;
    const res = await fetch(`${endpoint}/v1/chat/completions`, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        model: provider.model || 'gpt-3.5-turbo',
        max_tokens: 2500,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userMessage },
        ],
      }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `API error: ${res.status}`);
    }
    const data = await res.json();
    return data.choices?.[0]?.message?.content || '';
  }

  // Default: Anthropic
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 2500,
      system: systemPrompt,
      messages: [{ role: 'user', content: userMessage }],
    }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `API error: ${res.status}`);
  }
  const data = await res.json();
  return data.content?.[0]?.text || '';
}

function generateExportHtml(project) {
  const areasHtml = (project.selAreas || []).map(areaId => {
    const area = AREAS.find(a => a.id === areaId);
    const result = project.results?.[areaId] || '';
    const { body, flags } = splitFlags(result);
    const headingId = `area-${areaId}`;
    const escBody = body.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>');
    const escFlags = flags.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>');
    return `
    <section aria-labelledby="${headingId}" style="margin-bottom:2rem;padding:1.5rem;background:#fff;border:1px solid #ddd8f5;border-radius:8px;">
      <h2 id="${headingId}" style="color:#3a2b95;margin-bottom:1rem;">${area?.label || areaId}</h2>
      <div style="white-space:pre-wrap;line-height:1.7;">${escBody}</div>
      ${flags ? `
      <div role="region" aria-label="Flags for human review for ${area?.label || areaId}" style="margin-top:1.5rem;padding:1rem 1.25rem;background:#fff8e1;border:2px solid #ffca28;border-radius:6px;">
        <h3 style="color:#5d4037;margin-bottom:0.5rem;">⚑ FLAGS FOR HUMAN REVIEW</h3>
        <div style="white-space:pre-wrap;color:#5d4037;line-height:1.6;">${escFlags}</div>
      </div>` : ''}
    </section>`;
  }).join('\n');

  const escName = (project.name || 'Untitled').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const escDocName = (project.docName || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>My Expert Review Pal — ${escName}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:Arial,sans-serif;background:#f4f2fb;color:#1a1a2e;line-height:1.6;padding:2rem;}
  main{max-width:900px;margin:0 auto;}
  h1{color:#3a2b95;font-size:1.8rem;margin-bottom:0.5rem;}
  h2{font-size:1.4rem;}
  h3{font-size:1.1rem;}
  .meta{color:#555;margin-bottom:2rem;padding-bottom:1rem;border-bottom:2px solid #ddd8f5;}
  .meta span{margin-right:1.5rem;}
  @media print{body{background:#fff;padding:1rem;}section{break-inside:avoid;}}
  @media(max-width:600px){body{padding:0.75rem;}h1{font-size:1.4rem;}h2{font-size:1.2rem;}}
</style>
</head>
<body>
<main>
  <h1>My Expert Review Pal — ${escName}</h1>
  <div class="meta">
    <span><strong>Date:</strong> ${new Date(project.updatedAt || project.createdAt).toLocaleDateString()}</span>
    <span><strong>Format Style:</strong> ${project.fmtStyle || 'None'}</span>
    <span><strong>Areas Reviewed:</strong> ${(project.selAreas || []).map(id => AREAS.find(a=>a.id===id)?.label || id).join(', ')}</span>
    ${escDocName ? `<span><strong>Document:</strong> ${escDocName}</span>` : ''}
  </div>
  ${areasHtml}
</main>
</body>
</html>`;
}

function downloadHtml(html, filename) {
  const blob = new Blob([html], { type: 'text/html' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function parseMarkdownRuns(text) {
  const runs = [];
  const regex = /(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|__(.+?)__|_(.+?)_|`(.+?)`|([^*_`]+))/g;
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (match[2]) { runs.push(new TextRun({ text: match[2], bold: true, italics: true, font: 'Arial', size: 22 })); }
    else if (match[3]) { runs.push(new TextRun({ text: match[3], bold: true, font: 'Arial', size: 22 })); }
    else if (match[4]) { runs.push(new TextRun({ text: match[4], italics: true, font: 'Arial', size: 22 })); }
    else if (match[5]) { runs.push(new TextRun({ text: match[5], bold: true, font: 'Arial', size: 22 })); }
    else if (match[6]) { runs.push(new TextRun({ text: match[6], italics: true, font: 'Arial', size: 22 })); }
    else if (match[7]) { runs.push(new TextRun({ text: match[7], font: 'Courier New', size: 20 })); }
    else if (match[8]) { runs.push(new TextRun({ text: match[8], font: 'Arial', size: 22 })); }
  }
  return runs.length > 0 ? runs : [new TextRun({ text, font: 'Arial', size: 22 })];
}

async function generateAndDownloadDocx(project) {
  const children = [];

  children.push(new Paragraph({
    text: 'My Expert Review Pal — Review Report',
    heading: HeadingLevel.TITLE,
    spacing: { after: 200 },
  }));

  children.push(new Paragraph({
    children: [
      new TextRun({ text: 'Project: ', bold: true, font: 'Arial', size: 22 }),
      new TextRun({ text: project.name || 'Untitled', font: 'Arial', size: 22 }),
    ],
    spacing: { after: 80 },
  }));
  children.push(new Paragraph({
    children: [
      new TextRun({ text: 'Date: ', bold: true, font: 'Arial', size: 22 }),
      new TextRun({ text: new Date(project.updatedAt || project.createdAt).toLocaleDateString(), font: 'Arial', size: 22 }),
    ],
    spacing: { after: 80 },
  }));
  if (project.fmtStyle && project.fmtStyle !== 'None/Other') {
    children.push(new Paragraph({
      children: [
        new TextRun({ text: 'Format Style: ', bold: true, font: 'Arial', size: 22 }),
        new TextRun({ text: project.fmtStyle, font: 'Arial', size: 22 }),
      ],
      spacing: { after: 80 },
    }));
  }
  children.push(new Paragraph({
    children: [
      new TextRun({ text: 'Areas Reviewed: ', bold: true, font: 'Arial', size: 22 }),
      new TextRun({ text: (project.selAreas || []).map(id => AREAS.find(a => a.id === id)?.label || id).join(', '), font: 'Arial', size: 22 }),
    ],
    spacing: { after: 300 },
  }));

  children.push(new Paragraph({
    text: '',
    border: { bottom: { color: '3a2b95', space: 1, style: BorderStyle.SINGLE, size: 6 } },
    spacing: { after: 200 },
  }));

  for (const areaId of (project.selAreas || [])) {
    const area = AREAS.find(a => a.id === areaId);
    const result = project.results?.[areaId] || '';
    const { body, flags } = splitFlags(result);

    children.push(new Paragraph({
      text: area?.label || areaId,
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 300, after: 200 },
    }));

    for (const line of body.split('\n')) {
      if (!line.trim()) { children.push(new Paragraph({ text: '', spacing: { after: 80 } })); continue; }
      const trimmed = line.trim();
      if (trimmed.startsWith('## ')) {
        children.push(new Paragraph({ text: trimmed.slice(3), heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 100 } }));
      } else if (trimmed.startsWith('# ')) {
        children.push(new Paragraph({ text: trimmed.slice(2), heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }));
      } else if (trimmed.startsWith('- ') || trimmed.startsWith('• ')) {
        children.push(new Paragraph({ children: parseMarkdownRuns(trimmed.slice(2)), bullet: { level: 0 }, spacing: { after: 60 } }));
      } else {
        children.push(new Paragraph({ children: parseMarkdownRuns(trimmed), spacing: { after: 80 } }));
      }
    }

    if (flags) {
      children.push(new Paragraph({
        children: [new TextRun({ text: '⚑ FLAGS FOR HUMAN REVIEW', bold: true, color: '5d4037', font: 'Arial', size: 24 })],
        spacing: { before: 300, after: 120 },
        shading: { fill: 'FFF8E1' },
      }));
      for (const line of flags.split('\n')) {
        if (!line.trim()) continue;
        children.push(new Paragraph({
          children: parseMarkdownRuns(line.trim()),
          spacing: { after: 60 },
          shading: { fill: 'FFF8E1' },
        }));
      }
    }
  }

  const doc = new Document({
    creator: 'My Expert Review Pal',
    title: `Review Report — ${project.name || 'Untitled'}`,
    description: 'Accessible review report generated by My Expert Review Pal',
    styles: {
      default: {
        document: { run: { font: 'Arial', size: 22 } },
        heading1: { run: { font: 'Arial', size: 32, bold: true, color: '3a2b95' } },
        heading2: { run: { font: 'Arial', size: 28, bold: true, color: '3a2b95' } },
        heading3: { run: { font: 'Arial', size: 24, bold: true, color: '6f2fa6' } },
        title: { run: { font: 'Arial', size: 40, bold: true, color: '3a2b95' } },
      },
    },
    sections: [{
      properties: {
        page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
      },
      children,
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `${(project.name || 'review').replace(/\s+/g, '-')}-report.docx`);
}
async function generateChecklistDocx(project) {
  const areaCfg = getAreaCfg();
  const priorityOrder = { Critical: 0, High: 1, Medium: 2, Low: 3 };
  const priorityEmoji = { Critical: '🔴', High: '🟠', Medium: '🟡', Low: '🟢' };

  const areaItems = [];
  for (const areaId of (project.selAreas || [])) {
    const area = AREAS.find(a => a.id === areaId);
    const cfg = areaCfg[areaId] || {};
    const priority = cfg.priorityLevel || 'Medium';
    const result = project.results?.[areaId] || '';
    const { body, flags } = splitFlags(result);

    const items = [];
    const lines = body.split('\n');
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      // Extract lines that look like recommendations, issues, or action items
      if (/^\s*[-•*]\s+\*?\*?(Recommendation|Issue|Fix|Change|Replace|Add|Remove|Consider|Ensure|Verify|Check|Update|Revise|Correct)/i.test(trimmed) ||
          /\*\*Recommendation\*\*/i.test(trimmed) ||
          /\*\*Issue\*\*/i.test(trimmed)) {
        items.push(trimmed.replace(/^[-•*]\s+/, '').replace(/\*\*/g, ''));
      }
    }
    // Also pull from flags
    if (flags) {
      for (const line of flags.split('\n')) {
        const t = line.trim();
        if (t && t.length > 10 && !t.startsWith('#')) {
          items.push(t.replace(/^[-•*]\s+/, '').replace(/\*\*/g, ''));
        }
      }
    }
    // If no specific items extracted, grab lines that contain key action words
    if (items.length === 0) {
      for (const line of lines) {
        const t = line.trim();
        if (t && /should|must|need to|recommend|consider|ensure|missing|incorrect|error|fix/i.test(t) && t.length > 20 && t.length < 300) {
          items.push(t.replace(/^[-•*]\s+/, '').replace(/\*\*/g, ''));
          if (items.length >= 10) break;
        }
      }
    }

    if (items.length > 0) {
      areaItems.push({ area, priority, items });
    }
  }

  // Sort by priority
  areaItems.sort((a, b) => (priorityOrder[a.priority] || 2) - (priorityOrder[b.priority] || 2));

  const children = [];
  children.push(new Paragraph({
    text: 'Review Checklist',
    heading: HeadingLevel.TITLE,
    spacing: { after: 100 },
  }));
  children.push(new Paragraph({
    children: [
      new TextRun({ text: project.name || 'Untitled', bold: true, font: 'Arial', size: 22 }),
      new TextRun({ text: `  —  ${new Date().toLocaleDateString()}`, font: 'Arial', size: 22, color: '666666' }),
    ],
    spacing: { after: 200 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: 'This checklist contains the highest-priority items identified during review. Check off each item as it is addressed.', italics: true, font: 'Arial', size: 20, color: '666666' })],
    spacing: { after: 300 },
  }));

  for (const { area, priority, items } of areaItems) {
    children.push(new Paragraph({
      children: [
        new TextRun({ text: `${area?.label || 'Unknown'}`, bold: true, font: 'Arial', size: 26 }),
        new TextRun({ text: `  [${priority}]`, font: 'Arial', size: 22, color: priority === 'Critical' ? 'c62828' : priority === 'High' ? 'f9a825' : priority === 'Low' ? '888888' : '1565c0' }),
      ],
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 250, after: 120 },
    }));

    for (const item of items) {
      children.push(new Paragraph({
        children: [
          new TextRun({ text: '☐  ', font: 'Arial', size: 22 }),
          ...parseMarkdownRuns(item.length > 200 ? item.slice(0, 200) + '...' : item),
        ],
        spacing: { after: 80 },
        indent: { left: 360 },
      }));
    }
  }

  if (areaItems.length === 0) {
    children.push(new Paragraph({
      children: [new TextRun({ text: 'No specific action items were extracted from the review results.', italics: true, font: 'Arial', size: 22 })],
    }));
  }

  const doc = new Document({
    creator: 'My Expert Review Pal',
    title: `Review Checklist — ${project.name || 'Untitled'}`,
    description: 'Accessible review checklist generated by My Expert Review Pal',
    styles: {
      default: {
        document: { run: { font: 'Arial', size: 22 } },
        heading1: { run: { font: 'Arial', size: 26, bold: true, color: '3a2b95' } },
        title: { run: { font: 'Arial', size: 36, bold: true, color: '3a2b95' } },
      },
    },
    sections: [{
      properties: {
        page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
      },
      children,
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `${(project.name || 'review').replace(/\s+/g, '-')}-checklist.docx`);
}

const S = {
  sidebar: {
    width: 220, minWidth: 220, background: C.primary, color: '#fff', padding: '1.5rem 0',
    display: 'flex', flexDirection: 'column', position: 'sticky', top: 0, height: '100vh',
    overflowY: 'auto',
  },
  sidebarTitle: { fontSize: '1.15rem', fontWeight: 700, padding: '0 1.25rem 1.25rem', borderBottom: '1px solid rgba(255,255,255,0.15)' },
  navBtn: (active) => ({
    display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '0.7rem 1.25rem',
    background: active ? 'rgba(255,255,255,0.18)' : 'transparent', color: '#fff',
    border: 'none', cursor: 'pointer', fontSize: '0.95rem', textAlign: 'left',
    borderLeft: active ? '3px solid #fff' : '3px solid transparent',
    fontFamily: 'Arial, sans-serif',
  }),
  main: { flex: 1, padding: '2rem', maxWidth: 1100, minWidth: 0 },
  h1: { fontSize: '1.6rem', color: C.primary, marginBottom: '1.5rem' },
  h2: { fontSize: '1.3rem', color: C.primary, marginBottom: '1rem' },
  h3: { fontSize: '1.1rem', color: C.accent, marginBottom: '0.75rem' },
  card: {
    background: C.card, border: `1px solid ${C.border}`, borderRadius: 8,
    padding: '1.25rem', marginBottom: '1rem',
  },
  btn: (bg = C.primary, fg = '#fff') => ({
    background: bg, color: fg, border: 'none', borderRadius: 6, padding: '0.6rem 1.2rem',
    cursor: 'pointer', fontSize: '0.9rem', fontFamily: 'Arial, sans-serif', fontWeight: 600,
  }),
  btnOutline: {
    background: 'transparent', color: C.primary, border: `2px solid ${C.primary}`,
    borderRadius: 6, padding: '0.5rem 1rem', cursor: 'pointer', fontSize: '0.9rem',
    fontFamily: 'Arial, sans-serif', fontWeight: 600,
  },
  input: {
    width: '100%', padding: '0.6rem 0.8rem', border: `1px solid ${C.border}`,
    borderRadius: 6, fontSize: '0.95rem', fontFamily: 'Arial, sans-serif', color: C.text,
    background: '#fff',
  },
  textarea: {
    width: '100%', padding: '0.6rem 0.8rem', border: `1px solid ${C.border}`,
    borderRadius: 6, fontSize: '0.9rem', fontFamily: 'Arial, sans-serif', color: C.text,
    minHeight: 120, resize: 'vertical', background: '#fff',
  },
  badge: (bg, color = '#fff') => ({
    display: 'inline-block', background: bg, color, borderRadius: 12,
    padding: '0.15rem 0.6rem', fontSize: '0.75rem', fontWeight: 600,
  }),
  label: { display: 'block', marginBottom: '0.35rem', fontWeight: 600, fontSize: '0.9rem' },
  statsCard: {
    background: C.card, border: `1px solid ${C.border}`, borderRadius: 8,
    padding: '1rem 1.25rem', textAlign: 'center', flex: '1 1 140px',
  },
  toast: (type) => ({
    background: type === 'error' ? C.error : type === 'success' ? C.success : C.info,
    color: '#fff', padding: '0.75rem 1.25rem', borderRadius: 8, fontSize: '0.9rem',
    maxWidth: 380, boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
  }),
  flagBox: {
    background: C.flagBg, border: `2px solid ${C.flagBorder}`, borderRadius: 8,
    padding: '1rem 1.25rem', marginTop: '1rem',
  },
  flagTitle: { color: C.flagText, fontWeight: 700, marginBottom: '0.5rem', fontSize: '1rem' },
  flagBody: { color: C.flagText, whiteSpace: 'pre-wrap', lineHeight: 1.6, fontSize: '0.9rem' },
  tab: (active) => ({
    padding: '0.45rem 1rem', background: active ? C.primary : 'transparent',
    color: active ? '#fff' : C.primary, border: `1.5px solid ${C.primary}`, cursor: 'pointer',
    fontFamily: 'Arial, sans-serif', fontSize: '0.85rem', fontWeight: 600,
    borderRadius: 6, display: 'inline-flex', alignItems: 'center', whiteSpace: 'nowrap',
  }),
  stepBar: { display: 'flex', gap: 0, marginBottom: '2rem' },
  stepItem: (active, done) => ({
    flex: 1, textAlign: 'center', padding: '0.7rem 0.5rem',
    background: active ? C.primary : done ? C.accent : '#e0daf5',
    color: active || done ? '#fff' : C.text, fontWeight: 600, fontSize: '0.85rem',
    borderRight: '2px solid #fff',
  }),
  dropzone: (over) => ({
    border: `2px dashed ${over ? C.primary : C.border}`, borderRadius: 8,
    padding: '2.5rem', textAlign: 'center', cursor: 'pointer',
    background: over ? '#ebe7f7' : '#faf9ff', transition: 'background 0.2s',
  }),
  checkboxCard: (checked, disabled) => ({
    border: `2px solid ${checked ? C.primary : C.border}`,
    borderRadius: 8, padding: '0.85rem', cursor: disabled ? 'not-allowed' : 'pointer',
    background: disabled ? '#eee' : checked ? '#ede8fc' : '#fff',
    opacity: disabled ? 0.6 : 1, display: 'flex', gap: 10, alignItems: 'flex-start',
  }),
  branchNode: (depth) => ({
    marginLeft: depth * 24, padding: '0.75rem', marginBottom: '0.5rem',
    border: `1px solid ${C.border}`, borderRadius: 6,
    background: `rgba(111, 47, 166, ${Math.max(0.03, 0.12 - depth * 0.02)})`,
    color: '#000',
  }),
};

// ─── TOAST HOOK ─────────────────────────────────────────────────────────────────
function useToast() {
  const [toasts, setToasts] = useState([]);
  const add = useCallback((msg, type = 'success') => {
    const id = uid();
    setToasts(t => [...t, { id, msg, type }]);
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 3500);
  }, []);
  return { toasts, add };
}

// ─── CONFIRM DIALOG ────────────────────────────────────────────────────────────
function ConfirmDialog({ message, onConfirm, onCancel }) {
  return (
    <div role="alertdialog" aria-modal="true" aria-label="Confirmation dialog" style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex',
      alignItems: 'center', justifyContent: 'center', zIndex: 10000,
    }}>
      <div style={{ background: '#fff', borderRadius: 10, padding: '2rem', maxWidth: 400, width: '90%' }}>
        <p style={{ marginBottom: '1.25rem', fontSize: '1rem', lineHeight: 1.5 }}>{message}</p>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
          <button onClick={onCancel} style={S.btnOutline}>Cancel</button>
          <button onClick={onConfirm} style={S.btn(C.error)}>Delete</button>
        </div>
      </div>
    </div>
  );
}

// ─── BRANCH COMPONENT (RECURSIVE) ──────────────────────────────────────────────
function BranchNode({ branch, depth, onChange, onDelete, onAddChild }) {
  const [collapsed, setCollapsed] = useState(false);
  return (
    <div style={S.branchNode(depth)}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: collapsed ? 0 : 8 }}>
        <button onClick={() => setCollapsed(!collapsed)}
          style={{ ...S.btn('transparent', C.text), padding: '0.2rem 0.5rem', fontSize: '0.8rem', border: `1px solid ${C.border}` }}
          aria-expanded={!collapsed}
          aria-label={`${collapsed ? 'Expand' : 'Collapse'} branch ${branch.label || 'Unnamed'}`}>
          {collapsed ? '▶' : '▼'}
        </button>
        <span style={{ fontWeight: 600, fontSize: '0.85rem', color: '#000' }}>Depth {depth + 1}</span>
        <span style={{ fontSize: '0.85rem', color: '#000', flex: 1 }}>{branch.label || 'Unnamed branch'}</span>
        <button onClick={onDelete}
          style={{ ...S.btn(C.error), padding: '0.2rem 0.6rem', fontSize: '0.8rem' }}
          aria-label={`Delete branch ${branch.label || 'Unnamed'}`}>Delete</button>
      </div>
      {!collapsed && (
        <div>
          <label style={{ ...S.label, color: '#000' }}>Label
            <input style={{ ...S.input, marginTop: 4, color: '#000' }} value={branch.label || ''}
              onChange={e => onChange({ ...branch, label: e.target.value })} />
          </label>
          <label style={{ ...S.label, marginTop: 8, color: '#000' }}>Condition (When/If...)
            <input style={{ ...S.input, marginTop: 4, color: '#000' }} value={branch.condition || ''}
              placeholder="e.g., When the document contains statistical analyses..."
              onChange={e => onChange({ ...branch, condition: e.target.value })} />
          </label>
          <label style={{ ...S.label, marginTop: 8, color: '#000' }}>Additional Instructions
            <textarea style={{ ...S.textarea, minHeight: 70, marginTop: 4, color: '#000' }} value={branch.instructions || ''}
              onChange={e => onChange({ ...branch, instructions: e.target.value })} />
          </label>
          <label style={{ ...S.label, marginTop: 8, color: '#000' }}>Priority Level
            <select style={{ ...S.input, marginTop: 4, width: 'auto' }} value={branch.priority || 'Medium'}
              onChange={e => onChange({ ...branch, priority: e.target.value })}>
              {['Critical', 'High', 'Medium', 'Low'].map(p => <option key={p} value={p}>{p}</option>)}
            </select>
          </label>
          <button onClick={onAddChild}
            style={{ ...S.btn(C.accent), marginTop: 8, fontSize: '0.8rem', padding: '0.3rem 0.8rem' }}>
            + Add Sub-branch
          </button>
          {(branch.children || []).map((child, i) => (
            <BranchNode key={child.id || i} branch={child} depth={depth + 1}
              onChange={(updated) => {
                const nc = [...(branch.children || [])];
                nc[i] = updated;
                onChange({ ...branch, children: nc });
              }}
              onDelete={() => {
                onChange({ ...branch, children: (branch.children || []).filter((_, j) => j !== i) });
              }}
              onAddChild={() => {
                const nc = [...(branch.children || [])];
                nc[i] = {
                  ...nc[i],
                  children: [...(nc[i].children || []), { id: uid(), label: '', condition: '', instructions: '', children: [] }],
                };
                onChange({ ...branch, children: nc });
              }}
            />
          ))}
        </div>
      )}
    </div>
  );
}

// ─── DASHBOARD ──────────────────────────────────────────────────────────────────
function Dashboard({ navigate, toast }) {
  const [projects, setProjs] = useState([]);
  const [search, setSearch] = useState('');
  const [confirmDel, setConfirmDel] = useState(null);

  useEffect(() => {
    setProjs(getProjects().map(getProject).filter(Boolean));
  }, []);

  const rptCount = getReports().length;
  const completed = projects.filter(p => p.status === 'complete').length;
  const inProgress = projects.filter(p => p.status === 'draft').length;

  const filtered = projects.filter(p => {
    const q = search.toLowerCase();
    if (!q) return true;
    return (p.name || '').toLowerCase().includes(q) ||
      (p.notes || '').toLowerCase().includes(q) ||
      (p.tags || []).some(t => t.toLowerCase().includes(q));
  });

  const handleDelete = (id) => {
    deleteProject(id);
    setProjs(prev => prev.filter(p => p.id !== id));
    setConfirmDel(null);
    toast('Project deleted', 'success');
  };

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem', flexWrap: 'wrap', gap: 10 }}>
        <h1 style={S.h1} id="dash-heading"><BionicText>Dashboard</BionicText></h1>
        <button onClick={() => navigate('/review')} style={S.btn()} aria-label="Create new review">＋ New Review</button>
      </div>

      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: '1.5rem' }} role="group" aria-label="Project statistics">
        {[
          { label: 'Total Projects', val: projects.length },
          { label: 'Completed', val: completed },
          { label: 'In Progress', val: inProgress },
          { label: 'Reports Saved', val: rptCount },
        ].map(s => (
          <div key={s.label} style={S.statsCard}>
            <div style={{ fontSize: '1.5rem', fontWeight: 700, color: C.primary }}>{s.val}</div>
            <div style={{ fontSize: '0.8rem', color: '#666' }}><BionicText>{s.label}</BionicText></div>
          </div>
        ))}
      </div>

      <div style={{ marginBottom: '1.25rem' }}>
        <label htmlFor="dash-search" style={{ ...S.label, marginBottom: 6 }}>Search Projects</label>
        <input id="dash-search" style={S.input} placeholder="Search by name, tag, or notes..."
          value={search} onChange={e => setSearch(e.target.value)} />
      </div>

      {filtered.length === 0 && (
        <div style={{ ...S.card, textAlign: 'center', padding: '3rem' }}>
          <p style={{ fontSize: '1.1rem', marginBottom: '1rem', color: '#666' }}>
            <BionicText>{projects.length === 0 ? 'No projects yet. Start your first review!' : 'No projects match your search.'}</BionicText>
          </p>
          {projects.length === 0 && (
            <button onClick={() => navigate('/review')} style={S.btn()}>＋ Create Your First Review</button>
          )}
        </div>
      )}

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))', gap: 12 }}>
        {filtered.map(p => (
          <article key={p.id} style={S.card}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 8 }}>
              <h2 style={{ fontSize: '1.05rem', color: C.primary, margin: 0 }}><BionicText>{p.name || 'Untitled'}</BionicText></h2>
              <span style={S.badge(p.status === 'complete' ? C.success : C.warn, '#fff')}>
                {p.status === 'complete' ? 'Complete' : 'Draft'}
              </span>
            </div>
            <div style={{ fontSize: '0.8rem', color: '#888', marginBottom: 6 }}>
              {new Date(p.createdAt).toLocaleDateString()}
              {p.fmtStyle && p.fmtStyle !== 'None/Other' && (
                <span style={{ ...S.badge(C.accent), marginLeft: 8 }}>{p.fmtStyle}</span>
              )}
            </div>
            {p.notes && (
              <p style={{ fontSize: '0.85rem', color: '#555', marginBottom: 6 }}>
                <BionicText>{p.notes.slice(0, 100) + (p.notes.length > 100 ? '...' : '')}</BionicText>
              </p>
            )}
            <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap', marginBottom: 8 }}>
              {(p.tags || []).map(t => (
                <span key={t} style={S.badge('#e0daf5', C.accent)}>{t}</span>
              ))}
            </div>
            <div style={{ fontSize: '0.8rem', color: '#888', marginBottom: 10 }}>
              {(p.selAreas || []).length} review area{(p.selAreas || []).length !== 1 ? 's' : ''}
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <button onClick={() => navigate('/review', p.id)} style={S.btn()}>Open →</button>
              <button onClick={() => setConfirmDel(p.id)} style={S.btn(C.error)}
                aria-label={`Delete project ${p.name || 'Untitled'}`}>Delete</button>
            </div>
          </article>
        ))}
      </div>

      {confirmDel && (
        <ConfirmDialog message="Are you sure you want to delete this project? This cannot be undone."
          onConfirm={() => handleDelete(confirmDel)} onCancel={() => setConfirmDel(null)} />
      )}
    </div>
  );
}

// ─── REVIEW PAGE (3-STEP WIZARD) ────────────────────────────────────────────────
function ReviewPage({ projectId, navigate, toast }) {
  const [step, setStep] = useState(1);
  const [project, setProject] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [running, setRunning] = useState(false);
  const [areaStatus, setAreaStatus] = useState({});
  const [activeTab, setActiveTab] = useState(null);
  const fileInputRef = useRef(null);
  const areaCfg = getAreaCfg();
  const apiKey = getApiKey();

  useEffect(() => {
    if (projectId) {
      const p = getProject(projectId);
      if (p) {
        setProject(p);
        if (p.results && Object.keys(p.results).length > 0) {
          setStep(3);
          setActiveTab(p.selAreas?.[0] || null);
          const statuses = {};
          for (const a of p.selAreas || []) statuses[a] = 'done';
          setAreaStatus(statuses);
        } else if (p.docTxt) {
          setStep(2);
        }
        return;
      }
    }
    setProject({
      id: uid(), name: '', notes: '', tags: [], selAreas: [], fmtStyle: 'APA 7',
      docTxt: '', docName: '', extraCtx: '', status: 'draft',
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(), results: {},
    });
  }, [projectId]);

  const update = (changes) => setProject(prev => prev ? { ...prev, ...changes, updatedAt: new Date().toISOString() } : prev);

  const [fileLoading, setFileLoading] = useState(false);

  const handleFile = async (file) => {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['txt', 'md', 'csv', 'pdf', 'docx'].includes(ext)) {
      toast('Please upload .txt, .md, .csv, .pdf, or .docx files.', 'error');
      return;
    }
    if (ext === 'pdf') {
      setFileLoading(true);
      try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let text = '';
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const content = await page.getTextContent();
          text += content.items.map(item => item.str).join(' ') + '\n\n';
        }
        update({ docTxt: text.trim(), docName: file.name });
        toast(`Extracted ${pdf.numPages} pages from PDF.`, 'success');
      } catch (err) {
        toast(`PDF extraction failed: ${err.message}`, 'error');
      }
      setFileLoading(false);
    } else if (ext === 'docx') {
      setFileLoading(true);
      try {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        update({ docTxt: result.value, docName: file.name });
        toast('DOCX text extracted successfully.', 'success');
      } catch (err) {
        toast(`DOCX extraction failed: ${err.message}`, 'error');
      }
      setFileLoading(false);
    } else {
      const reader = new FileReader();
      reader.onload = (e) => update({ docTxt: e.target.result, docName: file.name });
      reader.readAsText(file);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault(); setDragOver(false);
    handleFile(e.dataTransfer.files[0]);
  };

  const toggleArea = (areaId) => {
    if (areaCfg[areaId]?.enabled === false) return;
    const sel = project.selAreas || [];
    update({ selAreas: sel.includes(areaId) ? sel.filter(a => a !== areaId) : [...sel, areaId] });
  };

  const saveToStorage = (p) => {
    const ids = getProjects();
    if (!ids.includes(p.id)) setProjects([...ids, p.id]);
    saveProjectData(p);
  };

  const saveDraft = () => {
    const p = { ...project, updatedAt: new Date().toISOString() };
    setProject(p);
    saveToStorage(p);
    toast('Draft saved!', 'success');
  };

  const runReview = async () => {
    const prov = getApiProvider();
    const needsKey = prov.type === 'anthropic' && !apiKey;
    const needsEndpoint = (prov.type === 'ollama' || prov.type === 'custom') && !prov.endpoint;
    if (needsKey) { toast('Please set your API key in API Settings first.', 'error'); return; }
    if (needsEndpoint) { toast('Please configure your API endpoint in API Settings first.', 'error'); return; }
    const activeAreas = (project.selAreas || []).filter(a => areaCfg[a]?.enabled !== false);
    if (activeAreas.length === 0) { toast('Please select at least one review area.', 'error'); return; }

    // Save project first
    const p = { ...project, updatedAt: new Date().toISOString() };
    saveToStorage(p);
    setProject(p);

    setStep(3);
    setRunning(true);
    const statuses = {};
    for (const a of activeAreas) statuses[a] = 'pending';
    setAreaStatus(statuses);
    setActiveTab(activeAreas[0]);
    const results = {};
    const docText = truncateText(p.docTxt, 14000);

    for (const areaId of activeAreas) {
      setAreaStatus(prev => ({ ...prev, [areaId]: 'running' }));
      try {
        const sysPrompt = buildSystemPrompt(areaId, areaCfg, p.fmtStyle, p.extraCtx);
        const userMsg = `Please review the following document:\n\n${docText}`;
        const result = await callApi(apiKey, sysPrompt, userMsg);
        results[areaId] = result;
        setAreaStatus(prev => ({ ...prev, [areaId]: 'done' }));
      } catch (err) {
        results[areaId] = `Error: ${err.message}`;
        setAreaStatus(prev => ({ ...prev, [areaId]: 'error' }));
        toast(`Error reviewing ${AREAS.find(a=>a.id===areaId)?.label}: ${err.message}`, 'error');
      }
      // Save progress incrementally
      const updatedProject = { ...p, results: { ...results }, status: Object.keys(results).length === activeAreas.length ? 'complete' : 'draft' };
      setProject(updatedProject);
      saveToStorage(updatedProject);
    }

    const finalProject = { ...p, results, status: 'complete', updatedAt: new Date().toISOString() };
    setProject(finalProject);
    saveToStorage(finalProject);
    setRunning(false);
    toast('Review complete!', 'success');
  };

  const exportHtml = () => {
    const html = generateExportHtml(project);
    downloadHtml(html, `${(project.name || 'review').replace(/\s+/g, '-')}-report.html`);
    toast('HTML report exported!', 'success');
  };

  const saveReport = () => {
    const rpt = {
      id: uid(), projectId: project.id, projectName: project.name,
      fmtStyle: project.fmtStyle, selAreas: project.selAreas,
      results: project.results, createdAt: new Date().toISOString(),
    };
    saveReportData(rpt);
    const ids = getReports();
    if (!ids.includes(rpt.id)) setReports([...ids, rpt.id]);
    toast('Report saved!', 'success');
  };

  if (!project) return null;

  const activeAreas = (project.selAreas || []).filter(a => areaCfg[a]?.enabled !== false);
  const doneCount = Object.values(areaStatus).filter(s => s === 'done').length;
  const totalCount = (project.selAreas || []).length;
  const progressPct = totalCount > 0 ? (doneCount / totalCount) * 100 : 0;

  return (
    <div>
      <h1 style={S.h1}><BionicText>{projectId ? `Review: ${project.name || 'Untitled'}` : 'New Review'}</BionicText></h1>

      {/* Step indicator */}
      <div style={S.stepBar} role="group" aria-label="Review progress steps">
        {['1. Upload', '2. Configure', '3. Results'].map((label, i) => (
          <div key={i} style={S.stepItem(step === i + 1, step > i + 1)}
            role="status" aria-current={step === i + 1 ? 'step' : undefined}>
            {label}
          </div>
        ))}
      </div>

      {/* ── STEP 1: UPLOAD ── */}
      {step === 1 && (
        <div>
          <h2 style={S.h2}><BionicText>Upload or Paste Document</BionicText></h2>

          <div style={S.dropzone(dragOver)}
            onDragOver={e => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            role="button" tabIndex={0}
            aria-label="Upload document. Click or drag and drop a file."
            onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInputRef.current?.click(); } }}>
            <input ref={fileInputRef} type="file" accept=".txt,.md,.csv,.pdf,.docx"
              style={{ display: 'none' }} onChange={e => handleFile(e.target.files[0])} aria-hidden="true" />
            <div style={{ marginBottom: 8 }} aria-hidden="true"><Icons.upload size={36}/></div>
            <p style={{ fontWeight: 600, marginBottom: 4, color: C.text }}>Drop file here or click to upload</p>
            <p style={{ fontSize: '0.85rem', color: C.accent }}>Accepts .txt, .md, .csv, .pdf, .docx</p>
          </div>

          {fileLoading && (
            <div style={{ ...S.card, marginTop: 12, textAlign: 'center', padding: '1.5rem' }} role="status" aria-live="polite">
              <Icons.spinner size={24} style={{ animation: 'spin 1s linear infinite' }} />
              <p style={{ marginTop: 8, color: C.accent }}>Extracting text from file...</p>
              <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
            </div>
          )}

          {project.docName && (
            <p style={{ marginTop: 10, fontSize: '0.9rem', color: C.success, fontWeight: 600 }}>
              Loaded: {project.docName}
            </p>
          )}

          <div style={{ marginTop: 16 }}>
            <label htmlFor="paste-area" style={S.label}>Or paste text directly</label>
            <textarea id="paste-area" style={{ ...S.textarea, minHeight: 180 }}
              value={project.docTxt} onChange={e => update({ docTxt: e.target.value })}
              placeholder="Paste your document text here..." />
            <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 6, fontSize: '0.8rem', color: C.accent }}>
              <span>{(project.docTxt || '').length.toLocaleString()} characters</span>
              <span>~{wordCount(project.docTxt || '').toLocaleString()} words</span>
            </div>
          </div>

          <div style={{ marginTop: 16, textAlign: 'right' }}>
            <button onClick={() => {
              if (project.docTxt.trim()) setStep(2);
              else toast('Please add document text first.', 'error');
            }} style={S.btn()} disabled={!project.docTxt.trim()}>
              Continue →
            </button>
          </div>
        </div>
      )}

      {/* ── STEP 2: CONFIGURE ── */}
      {step === 2 && (
        <div>
          <h2 style={S.h2}><BionicText>Configure Review</BionicText></h2>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 16 }}>
            <div>
              <label htmlFor="prj-name" style={S.label}>Project Name *</label>
              <input id="prj-name" style={S.input} value={project.name}
                placeholder="e.g., Dissertation Chapter 3"
                onChange={e => update({ name: e.target.value })} />
            </div>
            <div>
              <label htmlFor="prj-fmt" style={S.label}>Format Style</label>
              <select id="prj-fmt" style={S.input} value={project.fmtStyle}
                onChange={e => update({ fmtStyle: e.target.value })}>
                {FMT_STYLES.map(f => <option key={f} value={f}>{f}</option>)}
              </select>
            </div>
          </div>

          <div style={{ marginBottom: 16 }}>
            <label htmlFor="prj-tags" style={S.label}>Tags (comma-separated)</label>
            <input id="prj-tags" style={S.input}
              defaultValue={(project.tags || []).join(', ')}
              placeholder="e.g., dissertation, chapter-3, behavior-analysis"
              onBlur={e => update({ tags: e.target.value.split(',').map(t => t.trim()).filter(Boolean) })}
              onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); update({ tags: e.target.value.split(',').map(t => t.trim()).filter(Boolean) }); } }} />
          </div>

          <div style={{ marginBottom: 16 }}>
            <label htmlFor="prj-notes" style={S.label}>Notes</label>
            <textarea id="prj-notes" style={{ ...S.textarea, minHeight: 70 }} value={project.notes}
              placeholder="Any notes about this project..."
              onChange={e => update({ notes: e.target.value })} />
          </div>

          <h3 style={S.h3}>Select Review Areas</h3>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(250px, 1fr))', gap: 10, marginBottom: 16 }}>
            {AREAS.map(area => {
              const disabled = areaCfg[area.id]?.enabled === false;
              const checked = (project.selAreas || []).includes(area.id);
              return (
                <div key={area.id} style={S.checkboxCard(checked, disabled)}
                  onClick={() => !disabled && toggleArea(area.id)}
                  role="checkbox" aria-checked={checked} aria-disabled={disabled} tabIndex={0}
                  aria-label={`${area.label}: ${area.description}${disabled ? ' (Disabled in Configure)' : ''}`}
                  onKeyDown={e => { if (e.key === ' ' || e.key === 'Enter') { e.preventDefault(); if (!disabled) toggleArea(area.id); } }}>
                  <input type="checkbox" checked={checked} disabled={disabled} readOnly
                    tabIndex={-1} style={{ marginTop: 2, accentColor: C.primary }} aria-hidden="true" />
                  <div>
                    <div style={{ fontWeight: 600, fontSize: '0.9rem', display: 'flex', alignItems: 'center', gap: 6 }}><AreaIcon id={area.id} size={16}/> {area.label}</div>
                    <div style={{ fontSize: '0.8rem', color: C.accent, marginTop: 2 }}>{area.description}</div>
                    {disabled && <div style={{ fontSize: '0.75rem', color: C.error, marginTop: 4, fontWeight: 600 }}>Disabled in Configure</div>}
                  </div>
                </div>
              );
            })}
          </div>

          <div style={{ marginBottom: 16 }}>
            <label htmlFor="prj-ctx" style={S.label}>Additional Context (optional)</label>
            <textarea id="prj-ctx" style={{ ...S.textarea, minHeight: 60 }} value={project.extraCtx || ''}
              placeholder="e.g., This is a 5th-year doctoral dissertation in behavior analysis"
              onChange={e => update({ extraCtx: e.target.value })} />
          </div>

          <div style={{ display: 'flex', gap: 10, justifyContent: 'space-between', flexWrap: 'wrap' }}>
            <button onClick={() => setStep(1)} style={S.btnOutline}>← Back</button>
            <div style={{ display: 'flex', gap: 10 }}>
              <button onClick={saveDraft} style={S.btnOutline}>Save Draft</button>
              <button onClick={runReview}
                style={S.btn((() => { const pv = getApiProvider(); const ok = (pv.type === 'anthropic' && apiKey) || (pv.type !== 'anthropic' && pv.endpoint); return ok && activeAreas.length > 0 ? C.primary : '#999'; })())}
                disabled={(() => { const pv = getApiProvider(); const ok = (pv.type === 'anthropic' && apiKey) || (pv.type !== 'anthropic' && pv.endpoint); return !ok || activeAreas.length === 0; })()}>
                Run Review ({activeAreas.length} area{activeAreas.length !== 1 ? 's' : ''})
              </button>
            </div>
          </div>
          {!apiKey && getApiProvider().type === 'anthropic' && (
            <p style={{ marginTop: 8, fontSize: '0.85rem', color: C.error }}>
              Set your API key in API Settings to run reviews.
            </p>
          )}
        </div>
      )}

      {/* ── STEP 3: RESULTS ── */}
      {step === 3 && (
        <div>
          <h2 style={S.h2}><BionicText>Review Results</BionicText></h2>

          {/* Progress bar */}
          {running && (
            <div style={{ marginBottom: 16 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6, fontSize: '0.85rem' }}>
                <span>Progress</span>
                <span>{doneCount} / {totalCount} areas complete</span>
              </div>
              <div style={{ background: '#e0daf5', borderRadius: 6, height: 10, overflow: 'hidden' }}
                role="progressbar" aria-valuenow={doneCount} aria-valuemin={0} aria-valuemax={totalCount}
                aria-label={`Review progress: ${doneCount} of ${totalCount} areas complete`}>
                <div style={{ background: C.primary, height: '100%', width: `${progressPct}%`, transition: 'width 0.3s' }} />
              </div>
              <div style={{ display: 'flex', gap: 8, marginTop: 10, flexWrap: 'wrap' }}>
                {(project.selAreas || []).map(aId => {
                  const st = areaStatus[aId];
                  const area = AREAS.find(a => a.id === aId);
                  return (
                    <span key={aId} style={{ fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: 4 }}>
                      {st === 'running' ? <Icons.clock size={14}/> : st === 'done' ? <Icons.check size={14} style={{color:C.success}}/> : st === 'error' ? <Icons.x size={14} style={{color:C.error}}/> : <span style={{width:14,height:14,display:'inline-block',border:`1px solid ${C.border}`,borderRadius:2}}/>}
                      {area?.label}
                    </span>
                  );
                })}
              </div>
            </div>
          )}

          {/* Result tabs */}
          {(project.selAreas || []).length > 0 && (
            <div role="tablist" aria-label="Review area results"
              style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 16, paddingBottom: 12, borderBottom: `2px solid ${C.border}` }}>
              {(project.selAreas || []).map((aId) => {
                const area = AREAS.find(a => a.id === aId);
                const st = areaStatus[aId];
                return (
                  <button key={aId} role="tab" aria-selected={activeTab === aId}
                    aria-controls={`tabpanel-${aId}`} id={`tab-${aId}`}
                    onClick={() => setActiveTab(aId)}
                    style={S.tab(activeTab === aId)}>
                    {st === 'running' ? <Icons.clock size={14} style={{marginRight:4}}/> : st === 'done' ? <Icons.check size={14} style={{color:C.success,marginRight:4}}/> : st === 'error' ? <Icons.x size={14} style={{color:C.error,marginRight:4}}/> : null}
                    <AreaIcon id={aId} size={14} style={{marginRight:4}}/> {area?.label}
                  </button>
                );
              })}
            </div>
          )}

          {/* Tab panel content */}
          {activeTab && project.results?.[activeTab] && (() => {
            const { body, flags } = splitFlags(project.results[activeTab]);
            const area = AREAS.find(a => a.id === activeTab);
            return (
              <div role="tabpanel" id={`tabpanel-${activeTab}`} aria-labelledby={`tab-${activeTab}`}>
                <div style={S.card}>
                  <div style={{ whiteSpace: 'pre-wrap', lineHeight: 1.7, fontSize: '0.9rem' }}><BionicText>{body}</BionicText></div>
                </div>
                {flags && (
                  <div style={S.flagBox} role="region" aria-label={`Flags for human review for ${area?.label}`}>
                    <div style={S.flagTitle}>⚑ FLAGS FOR HUMAN REVIEW</div>
                    <div style={S.flagBody}><BionicText>{flags}</BionicText></div>
                  </div>
                )}
              </div>
            );
          })()}

          {activeTab && areaStatus[activeTab] === 'running' && (
            <div style={S.card} role="status" aria-live="polite">
              <p style={{ color: '#888', fontStyle: 'italic' }}>Analyzing... please wait.</p>
            </div>
          )}

          {/* Action buttons */}
          <div style={{ display: 'flex', gap: 10, marginTop: 20, flexWrap: 'wrap' }}>
            <button onClick={() => setStep(2)} style={S.btnOutline}>← Back to Configure</button>
            <button onClick={saveDraft} style={S.btnOutline}>Save Draft</button>
            {!running && Object.keys(project.results || {}).length > 0 && (
              <>
                <button onClick={exportHtml} style={S.btn(C.secondary)}>Export HTML</button>
                <button onClick={async () => { await generateAndDownloadDocx(project); toast('DOCX report downloaded!', 'success'); }} style={S.btn(C.info)}>Download DOCX</button>
                <button onClick={async () => { await generateChecklistDocx(project); toast('Checklist downloaded!', 'success'); }} style={S.btn(C.success)}>Download Checklist</button>
                <button onClick={saveReport} style={S.btn(C.accent)}>Save Report</button>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── CONFIGURE AREAS ────────────────────────────────────────────────────────────
function ConfigureAreas({ toast }) {
  const [cfg, setCfg] = useState(getAreaCfg);
  const [selected, setSelected] = useState('writing');
  const [showTips, setShowTips] = useState(false);
  const [combinedPrompt, setCombinedPrompt] = useState('');
  const fileRef = useRef(null);
  const keFileRef = useRef(null);

  const area = AREAS.find(a => a.id === selected);
  const ac = cfg[selected] || { enabled: true, prompt: DEFAULT_PROMPTS[selected], files: [], branches: [] };

  const updateArea = (changes) => {
    setCfg(prev => ({ ...prev, [selected]: { ...ac, ...changes } }));
  };

  const isCustomized = (areaId) => {
    const c = cfg[areaId] || {};
    return (c.prompt && c.prompt !== DEFAULT_PROMPTS[areaId]) ||
      (c.files || []).length > 0 || (c.branches || []).length > 0 || (c.knowledgeExamples || []).length > 0;
  };

  const handleFileUpload = async (file) => {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['txt', 'md', 'csv', 'pdf', 'docx'].includes(ext)) { toast('Only .txt, .md, .csv, .pdf, .docx files accepted.', 'error'); return; }
    if (ext === 'pdf') {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let text = '';
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const content = await page.getTextContent();
          text += content.items.map(item => item.str).join(' ') + '\n\n';
        }
        updateArea({ files: [...(ac.files || []), { name: file.name, size: file.size, content: text.trim() }] });
        toast('PDF reference file added.', 'success');
      } catch (err) { toast(`PDF extraction failed: ${err.message}`, 'error'); }
    } else if (ext === 'docx') {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        updateArea({ files: [...(ac.files || []), { name: file.name, size: file.size, content: result.value }] });
        toast('DOCX reference file added.', 'success');
      } catch (err) { toast(`DOCX extraction failed: ${err.message}`, 'error'); }
    } else {
      const reader = new FileReader();
      reader.onload = (e) => {
        updateArea({ files: [...(ac.files || []), { name: file.name, size: file.size, content: e.target.result }] });
      };
      reader.readAsText(file);
    }
  };

  const addBranch = () => {
    updateArea({
      branches: [...(ac.branches || []), { id: uid(), label: '', condition: '', instructions: '', children: [] }],
    });
  };

  const saveAll = () => {
    setAreaCfg(cfg);
    toast('Configuration saved!', 'success');
  };

  return (
    <div>
      <h1 style={S.h1}><BionicText>Configure Review Areas</BionicText></h1>
      <div style={{ display: 'grid', gridTemplateColumns: '220px 1fr', gap: 20 }}>
        {/* Left: area list */}
        <nav aria-label="Review areas list">
          {AREAS.map(a => (
            <button key={a.id} onClick={() => { setSelected(a.id); setShowTips(false); }}
              style={{
                display: 'flex', alignItems: 'center', gap: 8, width: '100%',
                padding: '0.6rem 0.8rem', marginBottom: 6,
                color: selected === a.id ? '#fff' : C.text,
                background: selected === a.id ? C.primary : '#fff',
                border: `1px solid ${C.border}`, borderRadius: 6,
                cursor: 'pointer', fontSize: '0.9rem', textAlign: 'left',
                fontFamily: 'Arial, sans-serif',
              }}
              aria-current={selected === a.id ? 'true' : undefined}>
              <span aria-hidden="true"><AreaIcon id={a.id} size={16}/></span>
              <span style={{ flex: 1 }}>{a.label}</span>
              <span style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
                {isCustomized(a.id) && (
                  <span title="Customized" aria-label="Customized"
                    style={{ color: selected === a.id ? '#fff' : C.accent, fontSize: '0.9rem' }}>●</span>
                )}
                <span style={{ ...S.badge(cfg[a.id]?.enabled === false ? C.error : C.success), fontSize: '0.65rem' }}>
                  {cfg[a.id]?.enabled === false ? 'OFF' : 'ON'}
                </span>
              </span>
            </button>
          ))}
          <button onClick={() => {
            let master = '# MASTER PROMPT — All Review Areas\n\n';
            for (const a of AREAS) {
              const c = cfg[a.id];
              if (c?.enabled === false) continue;
              master += `=== ${a.label} ===\n`;
              master += (c?.prompt || DEFAULT_PROMPTS[a.id]) + '\n';
              if (c?.branches?.length > 0) master += '\n--- BRANCHES ---' + flattenBranches(c.branches) + '\n';
              if (c?.files?.length > 0) { for (const f of c.files) master += `\n--- REF: ${f.name} ---\n${truncateText(f.content, 1000)}\n`; }
              if (c?.knowledgeExamples?.length > 0) { for (const ex of c.knowledgeExamples) master += `\n--- EXAMPLE (${ex.sentiment === 'liked' ? 'GOOD' : 'BAD'}): ${ex.name} ---\n${truncateText(ex.content, 500)}\n${ex.notes ? 'Notes: ' + ex.notes + '\n' : ''}`; }
              master += '\n\n';
            }
            setCombinedPrompt(master);
            toast('Master prompt generated from all enabled areas.', 'success');
          }}
            style={{
              display: 'flex', alignItems: 'center', gap: 8, width: '100%',
              padding: '0.6rem 0.8rem', marginTop: 12,
              color: '#fff', background: C.accent,
              border: `1px solid ${C.accent}`, borderRadius: 6,
              cursor: 'pointer', fontSize: '0.85rem', textAlign: 'left',
              fontFamily: 'Arial, sans-serif', fontWeight: 600,
            }}>
            <Icons.clip size={16}/> Generate Master Prompt
          </button>
        </nav>

        {/* Right: config panel */}
        <div style={S.card}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
            <div>
              <h2 style={{ ...S.h2, marginBottom: 4, display: 'flex', alignItems: 'center', gap: 8 }}><AreaIcon id={selected} size={20}/> <BionicText>{area?.label}</BionicText></h2>
              <p style={{ fontSize: '0.85rem', color: '#666' }}><BionicText>{area?.description}</BionicText></p>
            </div>
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer' }}>
              <input type="checkbox" checked={ac.enabled !== false}
                onChange={e => updateArea({ enabled: e.target.checked })}
                style={{ width: 18, height: 18, accentColor: C.primary }}
                aria-label={`Enable ${area?.label} review area`} />
              <span style={{ fontWeight: 600, fontSize: '0.9rem' }}>Enabled</span>
            </label>
          </div>

          {/* Location Marking */}
          <div style={{ marginBottom: 16 }}>
            <h3 style={S.h3}>How do you want to mark areas for improvement? <span style={{ fontWeight: 400, fontSize: '0.8rem', color: '#888' }}>(optional)</span></h3>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
              {[
                ['Page number', 'Identify by page number (e.g., Page 3)'],
                ['Paragraph number', 'Identify by paragraph count (e.g., Paragraph 12)'],
                ['Sentence with first 5 words', 'Quote the first 5 words of the sentence'],
                ['Section heading', 'Identify by the nearest section heading'],
                ['Line number', 'Identify by approximate line number'],
                ['Highlighted quote', 'Quote the exact problematic text'],
              ].map(([label, tip]) => {
                const checked = (ac.locationMarking || []).includes(label);
                return (
                  <label key={label} title={tip} style={{
                    display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: '0.85rem',
                    padding: '0.35rem 0.7rem', borderRadius: 5,
                    border: `1.5px solid ${checked ? C.primary : C.border}`,
                    background: checked ? '#ede8fc' : 'transparent',
                  }}>
                    <input type="checkbox" checked={checked}
                      onChange={() => {
                        const cur = ac.locationMarking || [];
                        updateArea({ locationMarking: checked ? cur.filter(l => l !== label) : [...cur, label] });
                      }}
                      style={{ accentColor: C.primary, width: 16, height: 16 }} />
                    {label}
                  </label>
                );
              })}
            </div>
          </div>

          {/* Priority Level */}
          <div style={{ marginBottom: 16 }}>
            <h3 style={S.h3}>Priority Level <span style={{ fontWeight: 400, fontSize: '0.8rem', color: '#888' }}>(optional)</span></h3>
            <p style={{ fontSize: '0.85rem', color: '#666', marginBottom: 8 }}>
              Set the importance of this review area. Higher priority areas get more thorough analysis.
            </p>
            <div style={{ display: 'flex', gap: 8 }}>
              {[
                ['Critical', C.error],
                ['High', C.warn],
                ['Medium', C.info],
                ['Low', '#888'],
              ].map(([level, color]) => {
                const active = (ac.priorityLevel || 'Medium') === level;
                return (
                  <button key={level} onClick={() => updateArea({ priorityLevel: level })}
                    style={{
                      padding: '0.4rem 1rem', borderRadius: 6, cursor: 'pointer',
                      border: `2px solid ${active ? color : C.border}`,
                      background: active ? color : 'transparent',
                      color: active ? '#fff' : C.text,
                      fontFamily: 'Arial, sans-serif', fontSize: '0.85rem', fontWeight: active ? 700 : 400,
                    }}
                    aria-pressed={active}>
                    {level}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Prompt */}
          <h3 style={S.h3}>Prompt / Training Instructions</h3>
          <label htmlFor={`prompt-${selected}`} className="sr-only"
            style={{ position: 'absolute', width: 1, height: 1, overflow: 'hidden', clip: 'rect(0,0,0,0)' }}>
            Prompt instructions for {area?.label}
          </label>
          <textarea id={`prompt-${selected}`} style={{ ...S.textarea, minHeight: 200 }}
            value={ac.prompt || ''} onChange={e => updateArea({ prompt: e.target.value })}
            aria-label={`Prompt instructions for ${area?.label}`} />
          <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
            <button onClick={() => updateArea({ prompt: DEFAULT_PROMPTS[selected] })} style={S.btnOutline}>
              Reset to Default
            </button>
            <button onClick={() => setShowTips(!showTips)} style={S.btnOutline}
              aria-expanded={showTips}>
              {showTips ? 'Hide Tips' : 'Show Tips'}
            </button>
          </div>
          {showTips && (
            <div style={{ ...S.card, marginTop: 10, background: '#f0ecff', borderColor: C.accent }}>
              <p style={{ fontSize: '0.85rem', lineHeight: 1.6 }}>{AREA_TIPS[selected]}</p>
            </div>
          )}

          {/* Reference Files */}
          <h3 style={{ ...S.h3, marginTop: 20 }}>Reference Files</h3>
          <p style={{ fontSize: '0.85rem', color: '#666', marginBottom: 10 }}>
            Upload reference files (.txt, .md, .csv, .pdf, .docx). Their content will be included in the AI prompt for this area.
          </p>
          <button onClick={() => fileRef.current?.click()} style={S.btnOutline}>Upload File</button>
          <input ref={fileRef} type="file" accept=".txt,.md,.csv,.pdf,.docx" style={{ display: 'none' }}
            onChange={e => { handleFileUpload(e.target.files[0]); e.target.value = ''; }}
            aria-label="Upload reference file" />
          {(ac.files || []).length > 0 && (
            <ul style={{ marginTop: 10, listStyle: 'none', padding: 0 }}>
              {ac.files.map((f, i) => (
                <li key={i} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '0.4rem 0', borderBottom: `1px solid ${C.border}` }}>
                  <span style={{ flex: 1, fontSize: '0.85rem' }}><Icons.clip size={14} style={{flexShrink:0}}/> {f.name}</span>
                  <span style={{ fontSize: '0.75rem', color: '#888' }}>{(f.size / 1024).toFixed(1)} KB</span>
                  <button onClick={() => updateArea({ files: ac.files.filter((_, j) => j !== i) })}
                    style={{ ...S.btn(C.error), padding: '0.15rem 0.5rem', fontSize: '0.75rem' }}
                    aria-label={`Remove file ${f.name}`}>Remove</button>
                </li>
              ))}
            </ul>
          )}

          {/* Knowledge Examples */}
          <h3 style={{ ...S.h3, marginTop: 20 }}>Knowledge Examples</h3>
          <p style={{ fontSize: '0.85rem', color: '#666', marginBottom: 10 }}>
            Upload example files and mark them as good or bad examples. These help the AI understand your quality expectations.
          </p>
          <button onClick={() => keFileRef.current?.click()} style={S.btnOutline}>Upload Example</button>
          <input ref={keFileRef} type="file" accept=".txt,.md,.csv,.pdf,.docx" style={{ display: 'none' }}
            onChange={async (e) => {
              const file = e.target.files[0];
              if (!file) return;
              const ext = file.name.split('.').pop().toLowerCase();
              let content = '';
              if (ext === 'pdf') {
                try {
                  const ab = await file.arrayBuffer();
                  const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
                  for (let i = 1; i <= pdf.numPages; i++) { const pg = await pdf.getPage(i); const c = await pg.getTextContent(); content += c.items.map(it => it.str).join(' ') + '\n'; }
                } catch (err) { toast(`PDF error: ${err.message}`, 'error'); return; }
              } else if (ext === 'docx') {
                try { const ab = await file.arrayBuffer(); const r = await mammoth.extractRawText({ arrayBuffer: ab }); content = r.value; }
                catch (err) { toast(`DOCX error: ${err.message}`, 'error'); return; }
              } else {
                content = await new Promise((res) => { const r = new FileReader(); r.onload = (ev) => res(ev.target.result); r.readAsText(file); });
              }
              updateArea({ knowledgeExamples: [...(ac.knowledgeExamples || []), { name: file.name, size: file.size, content, notes: '', sentiment: 'liked' }] });
              e.target.value = '';
            }}
            aria-label="Upload knowledge example file" />
          {(ac.knowledgeExamples || []).length > 0 && (
            <div style={{ marginTop: 10 }}>
              {(ac.knowledgeExamples || []).map((ex, i) => (
                <div key={i} style={{ padding: '0.75rem', marginBottom: 8, border: `1px solid ${C.border}`, borderRadius: 6, background: '#faf9ff' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                    <span style={{ fontSize: '0.85rem', fontWeight: 600 }}><Icons.clip size={14}/> {ex.name}</span>
                    <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                      <span style={{ fontSize: '0.75rem', color: '#888' }}>{(ex.size / 1024).toFixed(1)} KB</span>
                      <button onClick={() => updateArea({ knowledgeExamples: (ac.knowledgeExamples || []).filter((_, j) => j !== i) })}
                        style={{ ...S.btn(C.error), padding: '0.15rem 0.5rem', fontSize: '0.75rem' }}
                        aria-label={`Remove example ${ex.name}`}>Remove</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', gap: 12, marginBottom: 8 }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 4, cursor: 'pointer', fontSize: '0.85rem' }}>
                      <input type="radio" name={`ke-sentiment-${i}`} checked={ex.sentiment === 'liked'}
                        onChange={() => { const kes = [...(ac.knowledgeExamples || [])]; kes[i] = { ...kes[i], sentiment: 'liked' }; updateArea({ knowledgeExamples: kes }); }}
                        style={{ accentColor: C.success }} />
                      Good Example
                    </label>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 4, cursor: 'pointer', fontSize: '0.85rem' }}>
                      <input type="radio" name={`ke-sentiment-${i}`} checked={ex.sentiment === 'disliked'}
                        onChange={() => { const kes = [...(ac.knowledgeExamples || [])]; kes[i] = { ...kes[i], sentiment: 'disliked' }; updateArea({ knowledgeExamples: kes }); }}
                        style={{ accentColor: C.error }} />
                      Bad Example
                    </label>
                  </div>
                  <textarea style={{ ...S.textarea, minHeight: 50, fontSize: '0.85rem' }}
                    placeholder="Why did you like/dislike this example?" value={ex.notes || ''}
                    onChange={(e) => { const kes = [...(ac.knowledgeExamples || [])]; kes[i] = { ...kes[i], notes: e.target.value }; updateArea({ knowledgeExamples: kes }); }} />
                </div>
              ))}
            </div>
          )}

          {/* Branching Logic */}
          <h3 style={{ ...S.h3, marginTop: 20 }}>Branching Logic</h3>
          <div style={{ ...S.card, background: '#f8f6ff', marginBottom: 12, borderColor: C.border, color: '#000' }}>
            <p style={{ fontSize: '0.85rem', lineHeight: 1.6, color: '#000' }}>
              <strong>Branches add conditional reasoning.</strong> When a condition is detected in the document, the AI applies additional specialized instructions. Branches can be nested infinitely for complex decision trees.
            </p>
            <p style={{ fontSize: '0.8rem', color: '#333', marginTop: 6 }}>
              Example: &quot;When the document contains statistical analyses&quot; → &quot;Pay special attention to effect sizes and confidence intervals.&quot;
            </p>
          </div>
          <button onClick={addBranch} style={S.btnOutline}>+ Add Branch</button>
          <div style={{ marginTop: 10 }}>
            {(ac.branches || []).map((branch, i) => (
              <BranchNode key={branch.id || i} branch={branch} depth={0}
                onChange={(updated) => {
                  const nb = [...(ac.branches || [])];
                  nb[i] = updated;
                  updateArea({ branches: nb });
                }}
                onDelete={() => updateArea({ branches: (ac.branches || []).filter((_, j) => j !== i) })}
                onAddChild={() => {
                  const nb = [...(ac.branches || [])];
                  nb[i] = {
                    ...nb[i],
                    children: [...(nb[i].children || []), { id: uid(), label: '', condition: '', instructions: '', children: [] }],
                  };
                  updateArea({ branches: nb });
                }}
              />
            ))}
          </div>

          <div style={{ marginTop: 24, display: 'flex', gap: 10, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
            <button onClick={() => {
              let combined = '';
              for (const a of AREAS) {
                const c = cfg[a.id];
                if (c?.enabled === false) continue;
                combined += `=== ${a.label} ===\n`;
                combined += (c?.prompt || DEFAULT_PROMPTS[a.id]) + '\n';
                if (c?.branches?.length > 0) combined += '\n--- BRANCHES ---' + flattenBranches(c.branches) + '\n';
                if (c?.files?.length > 0) { for (const f of c.files) combined += `\n--- REF: ${f.name} ---\n${truncateText(f.content, 1000)}\n`; }
                if (c?.knowledgeExamples?.length > 0) { for (const ex of c.knowledgeExamples) combined += `\n--- EXAMPLE (${ex.sentiment === 'liked' ? 'GOOD' : 'BAD'}): ${ex.name} ---\n${truncateText(ex.content, 500)}\n${ex.notes ? 'Notes: ' + ex.notes + '\n' : ''}`; }
                combined += '\n\n';
              }
              setCombinedPrompt(combined);
            }} style={S.btnOutline}>Generate Combined Prompt</button>
            <button onClick={() => {
              let md = `# My Expert Review Pal — Configuration Export\n\n**Date:** ${new Date().toLocaleDateString()}\n\n`;
              for (const a of AREAS) {
                const c = cfg[a.id];
                if (c?.enabled === false) continue;
                md += `## ${a.label}\n\n`;
                md += `**Status:** Enabled\n\n`;
                md += `### Prompt\n\n${c?.prompt || DEFAULT_PROMPTS[a.id]}\n\n`;
                if (c?.branches?.length > 0) {
                  md += `### Branches\n\n`;
                  const writeBranches = (branches, depth = 0) => {
                    for (const b of branches) {
                      md += `${'  '.repeat(depth)}- **${b.label || 'Unnamed'}**: When ${b.condition || '(no condition)'} → ${b.instructions || '(no instructions)'}\n`;
                      if (b.children?.length > 0) writeBranches(b.children, depth + 1);
                    }
                  };
                  writeBranches(c.branches);
                  md += '\n';
                }
                if (c?.files?.length > 0) { md += `### Reference Files\n\n`; for (const f of c.files) md += `- **${f.name}** (${(f.size/1024).toFixed(1)} KB)\n`; md += '\n'; }
                if (c?.knowledgeExamples?.length > 0) { md += `### Knowledge Examples\n\n`; for (const ex of c.knowledgeExamples) md += `- **${ex.name}** (${ex.sentiment === 'liked' ? 'Good' : 'Bad'})${ex.notes ? ': ' + ex.notes : ''}\n`; md += '\n'; }
                md += '---\n\n';
              }
              const blob = new Blob([md], { type: 'text/markdown' });
              const url = URL.createObjectURL(blob);
              const a = document.createElement('a'); a.href = url; a.download = `merp-config-${new Date().toISOString().slice(0,10)}.md`; a.click();
              URL.revokeObjectURL(url);
              toast('Markdown exported!', 'success');
            }} style={S.btnOutline}>Download as .md</button>
            <button onClick={saveAll} style={S.btn()}>Save All Changes</button>
          </div>

          {combinedPrompt && (
            <div style={{ marginTop: 16 }}>
              <h3 style={S.h3}>Combined Prompt</h3>
              <textarea readOnly value={combinedPrompt} style={{ ...S.textarea, minHeight: 200, fontFamily: 'monospace', fontSize: '0.8rem' }} />
              <button onClick={() => { navigator.clipboard.writeText(combinedPrompt); toast('Copied to clipboard!', 'success'); }}
                style={{ ...S.btnOutline, marginTop: 8 }}>Copy to Clipboard</button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── API SETTINGS ───────────────────────────────────────────────────────────────
function ApiSettings({ toast }) {
  const [key, setKey] = useState(getApiKey);
  const [show, setShow] = useState(false);
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState(null);
  const [provider, setProviderState] = useState(getApiProvider);

  const saveKey = () => { setApiKeyStorage(key); toast('API key saved!', 'success'); };

  const updateProvider = (changes) => {
    const updated = { ...provider, ...changes };
    setProviderState(updated);
    setApiProvider(updated);
  };

  const testConnection = async () => {
    setTesting(true); setTestResult(null);
    try {
      if (provider.type === 'ollama') {
        const endpoint = (provider.endpoint || 'http://localhost:11434').replace(/\/$/, '');
        const res = await fetch(`${endpoint}/api/tags`);
        if (res.ok) {
          const data = await res.json();
          const models = (data.models || []).map(m => m.name).join(', ');
          setTestResult({ ok: true, msg: `Connected! Available models: ${models || 'none found'}` });
        } else {
          setTestResult({ ok: false, msg: `Ollama error: ${res.status}` });
        }
      } else if (provider.type === 'custom') {
        const endpoint = (provider.endpoint || '').replace(/\/$/, '');
        const headers = { 'Content-Type': 'application/json' };
        if (provider.apiKey) headers['Authorization'] = `Bearer ${provider.apiKey}`;
        const res = await fetch(`${endpoint}/v1/models`, { headers });
        if (res.ok) {
          setTestResult({ ok: true, msg: 'Connection successful!' });
        } else {
          setTestResult({ ok: false, msg: `Error: ${res.status}` });
        }
      } else {
        const res = await fetch('https://api.anthropic.com/v1/messages', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': key,
            'anthropic-version': '2023-06-01',
            'anthropic-dangerous-direct-browser-access': 'true',
          },
          body: JSON.stringify({
            model: 'claude-sonnet-4-20250514',
            max_tokens: 50,
            messages: [{ role: 'user', content: 'Reply with just the word "connected".' }],
          }),
        });
        if (res.ok) {
          setTestResult({ ok: true, msg: 'Connection successful! API key is valid.' });
        } else {
          const err = await res.json().catch(() => ({}));
          setTestResult({ ok: false, msg: err.error?.message || `Error: ${res.status}` });
        }
      }
    } catch (e) {
      setTestResult({ ok: false, msg: `Network error: ${e.message}` });
    }
    setTesting(false);
  };

  const providerLabel = { anthropic: 'Anthropic Cloud', ollama: 'Ollama (Local)', custom: 'Custom Endpoint' };

  return (
    <div>
      <h1 style={S.h1}><BionicText>API Settings</BionicText></h1>

      {/* Provider selector */}
      <div style={{ ...S.card, maxWidth: 600, marginBottom: 16 }}>
        <h2 style={S.h2}><BionicText>Provider</BionicText></h2>
        <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
          {['anthropic', 'ollama', 'custom'].map(t => (
            <button key={t} onClick={() => { updateProvider({ type: t }); setTestResult(null); }}
              style={{
                ...S.btn(provider.type === t ? C.primary : 'transparent', provider.type === t ? '#fff' : C.text),
                border: `2px solid ${provider.type === t ? C.primary : C.border}`,
              }}>
              {providerLabel[t]}
            </button>
          ))}
        </div>
        <div style={{ ...S.badge(C.accent), marginBottom: 8 }}>Active: {providerLabel[provider.type]}</div>
      </div>

      <div style={{ ...S.card, maxWidth: 600 }}>
        {/* Anthropic Cloud */}
        {provider.type === 'anthropic' && (
          <>
            <h2 style={S.h2}><BionicText>Anthropic API Key</BionicText></h2>
            <label htmlFor="api-key-input" style={S.label}>API Key</label>
            <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
              <input id="api-key-input" type={show ? 'text' : 'password'} style={{ ...S.input, flex: 1 }}
                value={key} onChange={e => setKey(e.target.value)} placeholder="sk-ant-..." />
              <button onClick={() => setShow(!show)} style={S.btnOutline}
                aria-label={show ? 'Hide API key' : 'Show API key'}>
                {show ? 'Hide' : 'Show'}
              </button>
              <button onClick={saveKey} style={S.btn()}>Save</button>
            </div>
            <h3 style={S.h3}>Model</h3>
            <p style={{ fontSize: '0.9rem', marginBottom: 16, padding: '0.5rem 0.75rem', background: '#f0ecff', borderRadius: 6 }}>
              <code style={{ fontFamily: 'monospace' }}>claude-sonnet-4-20250514</code> (fixed)
            </p>
          </>
        )}

        {/* Ollama (Local) */}
        {provider.type === 'ollama' && (
          <>
            <h2 style={S.h2}><BionicText>Ollama (Local)</BionicText></h2>
            <label htmlFor="ollama-endpoint" style={S.label}>Endpoint URL</label>
            <input id="ollama-endpoint" style={{ ...S.input, marginBottom: 12 }}
              value={provider.endpoint || ''} onChange={e => updateProvider({ endpoint: e.target.value })}
              placeholder="http://localhost:11434" />
            <label htmlFor="ollama-model" style={S.label}>Model Name</label>
            <input id="ollama-model" style={{ ...S.input, marginBottom: 12 }}
              value={provider.model || ''} onChange={e => updateProvider({ model: e.target.value })}
              placeholder="e.g., llama3, mistral, codellama" />
            <p style={{ fontSize: '0.85rem', color: '#666', marginBottom: 12 }}>
              <BionicText>No API key needed for local Ollama. Make sure Ollama is running.</BionicText>
            </p>
          </>
        )}

        {/* Custom Endpoint */}
        {provider.type === 'custom' && (
          <>
            <h2 style={S.h2}><BionicText>Custom Endpoint (OpenAI-compatible)</BionicText></h2>
            <label htmlFor="custom-endpoint" style={S.label}>Endpoint URL</label>
            <input id="custom-endpoint" style={{ ...S.input, marginBottom: 12 }}
              value={provider.endpoint || ''} onChange={e => updateProvider({ endpoint: e.target.value })}
              placeholder="https://api.example.com" />
            <label htmlFor="custom-apikey" style={S.label}>API Key (optional)</label>
            <input id="custom-apikey" type="password" style={{ ...S.input, marginBottom: 12 }}
              value={provider.apiKey || ''} onChange={e => updateProvider({ apiKey: e.target.value })}
              placeholder="Bearer token..." />
            <label htmlFor="custom-model" style={S.label}>Model Name</label>
            <input id="custom-model" style={{ ...S.input, marginBottom: 12 }}
              value={provider.model || ''} onChange={e => updateProvider({ model: e.target.value })}
              placeholder="e.g., gpt-4, claude-3-opus" />
            <p style={{ fontSize: '0.85rem', color: '#666', marginBottom: 12 }}>
              Uses OpenAI-compatible format (POST to /v1/chat/completions).
            </p>
          </>
        )}

        <button onClick={testConnection} style={S.btn(C.accent)} disabled={testing ||
          (provider.type === 'anthropic' && !key)}>
          {testing ? 'Testing...' : 'Test Connection'}
        </button>
        {testResult && (
          <div style={{
            marginTop: 12, padding: '0.7rem', borderRadius: 6,
            background: testResult.ok ? '#e8f5e9' : '#fce4ec', fontSize: '0.9rem',
          }} role="alert">
            {testResult.ok ? <Icons.check size={16} style={{color:C.success,verticalAlign:'middle',marginRight:6}}/> : <Icons.x size={16} style={{color:C.error,verticalAlign:'middle',marginRight:6}}/>}{testResult.msg}
          </div>
        )}

        <div style={{ marginTop: 20, padding: '0.85rem', background: '#f8f6ff', borderRadius: 6, fontSize: '0.85rem', lineHeight: 1.6 }}>
          <strong>Privacy note:</strong> Your API key and settings are stored only in your browser&apos;s localStorage. They are never transmitted anywhere except directly to the configured API endpoint when running reviews.
        </div>
      </div>
    </div>
  );
}

// ─── REPORTS PAGE ───────────────────────────────────────────────────────────────
function ReportsPage({ toast }) {
  const [subTab, setSubTab] = useState('saved');
  const [reports, setRpts] = useState([]);
  const [expanded, setExpanded] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);

  // Pattern Analysis state
  const [filterText, setFilterText] = useState('');
  const [themes, setThemes] = useState('');
  const [analysisTypes, setAnalysisTypesState] = useState([]);
  const [analysisRunning, setAnalysisRunning] = useState(false);
  const [analysisResult, setAnalysisResult] = useState('');
  const [analyses, setAnalysesLocal] = useState([]);
  const [expandedAnalysis, setExpandedAnalysis] = useState(null);
  const [confirmDelAnalysis, setConfirmDelAnalysis] = useState(null);

  const analysisOptions = [
    'Most common error types',
    'Recurring flag themes',
    'Improvement over time',
    'Specific terminology patterns',
    'Strengths identified',
  ];

  const loadData = () => {
    setRpts(getReports().map(getReport).filter(Boolean));
    setAnalysesLocal(getAnalyses().map(getAnalysis).filter(Boolean));
  };

  useEffect(loadData, [subTab]);

  const handleDeleteReport = (id) => {
    deleteReport(id);
    setConfirmDel(null);
    loadData();
    toast('Report deleted.', 'success');
  };

  const handleDeleteAnalysis = (id) => {
    deleteAnalysis(id);
    setConfirmDelAnalysis(null);
    loadData();
    toast('Analysis deleted.', 'success');
  };

  const exportReport = (rpt) => {
    const fakeProject = {
      name: rpt.projectName, fmtStyle: rpt.fmtStyle, selAreas: rpt.selAreas,
      results: rpt.results, createdAt: rpt.createdAt, updatedAt: rpt.createdAt,
    };
    const html = generateExportHtml(fakeProject);
    downloadHtml(html, `${(rpt.projectName || 'report').replace(/\s+/g, '-')}-report.html`);
    toast('Report exported!', 'success');
  };

  const toggleType = (t) => {
    setAnalysisTypesState(prev => prev.includes(t) ? prev.filter(x => x !== t) : [...prev, t]);
  };

  const runAnalysis = async () => {
    const apiKey = getApiKey();
    if (!apiKey) { toast('Set API key first.', 'error'); return; }
    if (analysisTypes.length === 0) { toast('Select at least one analysis type.', 'error'); return; }

    setAnalysisRunning(true);
    setAnalysisResult('');

    const projects = getProjects().map(getProject).filter(Boolean);
    const q = filterText.toLowerCase();
    const matching = q ? projects.filter(p =>
      (p.name || '').toLowerCase().includes(q) ||
      (p.tags || []).some(t => t.toLowerCase().includes(q))
    ) : projects;

    if (matching.length === 0) {
      toast('No matching projects found.', 'error');
      setAnalysisRunning(false);
      return;
    }

    let compiled = `PROJECTS ANALYZED (${matching.length} total):\n\n`;
    for (const p of matching) {
      compiled += `--- Project: ${p.name} (${new Date(p.createdAt).toLocaleDateString()}) ---\n`;
      compiled += `Tags: ${(p.tags || []).join(', ')}\nNotes: ${p.notes || 'None'}\n`;
      if (p.results) {
        for (const [areaId, result] of Object.entries(p.results)) {
          const area = AREAS.find(a => a.id === areaId);
          compiled += `\n[${area?.label || areaId}]:\n${truncateText(result, 2000)}\n`;
        }
      }
      compiled += '\n';
    }

    const sysPrompt = `You are an expert academic review analyst. Analyze the following compiled review data from multiple academic document reviews.

Requested analysis types: ${analysisTypes.join(', ')}
${themes ? `Specific words/themes to look for: ${themes}` : ''}

Provide a thorough, well-organized analysis with clear headings and actionable insights. If analyzing improvement over time, use project dates to establish chronology. Be specific with examples from the data.`;

    try {
      const result = await callApi(apiKey, sysPrompt, compiled);
      setAnalysisResult(result);
    } catch (e) {
      toast(`Analysis error: ${e.message}`, 'error');
      setAnalysisResult(`Error: ${e.message}`);
    }
    setAnalysisRunning(false);
  };

  const saveAnalysisReport = () => {
    const a = {
      id: uid(), name: `Analysis — ${new Date().toLocaleDateString()}`,
      filterText, themes, analysisTypes, result: analysisResult,
      createdAt: new Date().toISOString(),
    };
    saveAnalysisData(a);
    const ids = getAnalyses();
    setAnalyses([...ids, a.id]);
    loadData();
    toast('Analysis saved!', 'success');
  };

  return (
    <div>
      <h1 style={S.h1}><BionicText>Reports</BionicText></h1>

      {/* Sub-tabs */}
      <div role="tablist" aria-label="Report sections" style={{ display: 'flex', gap: 6, marginBottom: 20 }}>
        <button role="tab" aria-selected={subTab === 'saved'} id="tab-saved"
          aria-controls="panel-saved" onClick={() => setSubTab('saved')}
          style={S.tab(subTab === 'saved')}>
          Saved Reports
        </button>
        <button role="tab" aria-selected={subTab === 'analysis'} id="tab-analysis"
          aria-controls="panel-analysis" onClick={() => setSubTab('analysis')}
          style={S.tab(subTab === 'analysis')}>
          Pattern Analysis
        </button>
      </div>

      {/* ── SAVED REPORTS ── */}
      {subTab === 'saved' && (
        <div role="tabpanel" id="panel-saved" aria-labelledby="tab-saved">
          {reports.length === 0 ? (
            <div style={{ ...S.card, textAlign: 'center', padding: '2rem' }}>
              <p style={{ color: '#666' }}>No saved reports yet. Complete a review and save the report to see it here.</p>
            </div>
          ) : (
            reports.map(rpt => (
              <article key={rpt.id} style={S.card}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
                  <div>
                    <h2 style={{ fontSize: '1rem', color: C.primary, margin: 0 }}><BionicText>{rpt.projectName || 'Untitled'}</BionicText></h2>
                    <div style={{ fontSize: '0.8rem', color: '#888', marginTop: 4 }}>
                      {new Date(rpt.createdAt).toLocaleDateString()}
                      {rpt.fmtStyle && <span style={{ ...S.badge(C.accent), marginLeft: 8 }}>{rpt.fmtStyle}</span>}
                      <span style={{ marginLeft: 8 }}>
                        {(rpt.selAreas || []).map(id => AREAS.find(a=>a.id===id)?.label || id).join(', ')}
                      </span>
                    </div>
                  </div>
                  <div style={{ display: 'flex', gap: 8 }}>
                    <button onClick={() => setExpanded(expanded === rpt.id ? null : rpt.id)}
                      style={S.btnOutline} aria-expanded={expanded === rpt.id}>
                      {expanded === rpt.id ? 'Hide' : 'View'}
                    </button>
                    <button onClick={() => exportReport(rpt)} style={S.btn(C.secondary)}>Export HTML</button>
                    <button onClick={async () => {
                      const fakeProject = { name: rpt.projectName, fmtStyle: rpt.fmtStyle, selAreas: rpt.selAreas, results: rpt.results, createdAt: rpt.createdAt, updatedAt: rpt.createdAt };
                      await generateAndDownloadDocx(fakeProject);
                      toast('DOCX report downloaded!', 'success');
                    }} style={S.btn(C.info)}>DOCX</button>
                    <button onClick={async () => {
                      const fakeProject = { name: rpt.projectName, fmtStyle: rpt.fmtStyle, selAreas: rpt.selAreas, results: rpt.results, createdAt: rpt.createdAt, updatedAt: rpt.createdAt };
                      await generateChecklistDocx(fakeProject);
                      toast('Checklist downloaded!', 'success');
                    }} style={S.btn(C.success)}>Checklist</button>
                    <button onClick={() => setConfirmDel(rpt.id)} style={S.btn(C.error)}
                      aria-label={`Delete report ${rpt.projectName || 'Untitled'}`}>Delete</button>
                  </div>
                </div>

                {expanded === rpt.id && (
                  <div style={{ marginTop: 16 }}>
                    {(rpt.selAreas || []).map(aId => {
                      const area = AREAS.find(a => a.id === aId);
                      const { body, flags } = splitFlags(rpt.results?.[aId] || '');
                      return (
                        <details key={aId} style={{ marginBottom: 10, border: `1px solid ${C.border}`, borderRadius: 6 }}>
                          <summary style={{ padding: '0.7rem', cursor: 'pointer', fontWeight: 600, fontSize: '0.9rem', display: 'flex', alignItems: 'center', gap: 6 }}>
                            <AreaIcon id={aId} size={16}/> {area?.label}
                          </summary>
                          <div style={{ padding: '0.75rem' }}>
                            <div style={{ whiteSpace: 'pre-wrap', lineHeight: 1.7, fontSize: '0.85rem' }}><BionicText>{body}</BionicText></div>
                            {flags && (
                              <div style={S.flagBox}>
                                <div style={S.flagTitle}>⚑ FLAGS FOR HUMAN REVIEW</div>
                                <div style={S.flagBody}><BionicText>{flags}</BionicText></div>
                              </div>
                            )}
                          </div>
                        </details>
                      );
                    })}
                  </div>
                )}
              </article>
            ))
          )}
        </div>
      )}

      {/* ── PATTERN ANALYSIS ── */}
      {subTab === 'analysis' && (
        <div role="tabpanel" id="panel-analysis" aria-labelledby="tab-analysis">
          <div style={S.card}>
            <h2 style={S.h2}><BionicText>Analyze Patterns Across Reviews</BionicText></h2>

            <div style={{ marginBottom: 16 }}>
              <label htmlFor="analysis-filter" style={S.label}>Filter by person/tag (optional)</label>
              <input id="analysis-filter" style={S.input} value={filterText}
                onChange={e => setFilterText(e.target.value)}
                placeholder="Search project tags and names..." />
            </div>

            <div style={{ marginBottom: 16 }}>
              <label htmlFor="analysis-themes" style={S.label}>Specific words or themes to look for</label>
              <input id="analysis-themes" style={S.input} value={themes}
                onChange={e => setThemes(e.target.value)}
                placeholder="e.g., APA formatting, citations, passive voice" />
            </div>

            <fieldset style={{ border: 'none', marginBottom: 16, padding: 0 }}>
              <legend style={S.label}>Type of analysis</legend>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginTop: 4 }}>
                {analysisOptions.map(opt => (
                  <label key={opt} style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.9rem' }}>
                    <input type="checkbox" checked={analysisTypes.includes(opt)}
                      onChange={() => toggleType(opt)} style={{ accentColor: C.primary }} />
                    {opt}
                  </label>
                ))}
              </div>
            </fieldset>

            <button onClick={runAnalysis} style={S.btn()}
              disabled={analysisRunning || analysisTypes.length === 0}>
              {analysisRunning ? 'Analyzing...' : 'Generate Analysis'}
            </button>

            {analysisRunning && (
              <div style={{ marginTop: 16 }} role="status" aria-live="polite">
                <p style={{ fontStyle: 'italic', color: '#888' }}>Running pattern analysis... this may take a moment.</p>
              </div>
            )}

            {analysisResult && !analysisRunning && (
              <div style={{ marginTop: 16 }}>
                <h3 style={S.h3}>Analysis Results</h3>
                <div style={{ ...S.card, whiteSpace: 'pre-wrap', lineHeight: 1.7, fontSize: '0.9rem' }}>
                  <BionicText>{analysisResult}</BionicText>
                </div>
                <button onClick={saveAnalysisReport} style={{ ...S.btn(C.accent), marginTop: 10 }}>
                  Save Analysis Report
                </button>
              </div>
            )}
          </div>

          {/* Saved Analyses */}
          {analyses.length > 0 && (
            <div style={{ marginTop: 20 }}>
              <h3 style={S.h3}>Saved Analyses</h3>
              {analyses.map(a => (
                <article key={a.id} style={S.card}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div>
                      <h3 style={{ fontSize: '0.95rem', color: C.primary, margin: 0 }}>{a.name}</h3>
                      <div style={{ fontSize: '0.8rem', color: '#888', marginTop: 2 }}>
                        {new Date(a.createdAt).toLocaleDateString()} — {(a.analysisTypes || []).join(', ')}
                      </div>
                    </div>
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button onClick={() => setExpandedAnalysis(expandedAnalysis === a.id ? null : a.id)}
                        style={S.btnOutline} aria-expanded={expandedAnalysis === a.id}>
                        {expandedAnalysis === a.id ? 'Hide' : 'View'}
                      </button>
                      <button onClick={() => setConfirmDelAnalysis(a.id)} style={S.btn(C.error)}
                        aria-label={`Delete analysis ${a.name}`}>Delete</button>
                    </div>
                  </div>
                  {expandedAnalysis === a.id && (
                    <div style={{ marginTop: 12, whiteSpace: 'pre-wrap', lineHeight: 1.7, fontSize: '0.85rem' }}>
                      {a.result}
                    </div>
                  )}
                </article>
              ))}
            </div>
          )}
        </div>
      )}

      {confirmDel && (
        <ConfirmDialog message="Delete this report? This cannot be undone."
          onConfirm={() => handleDeleteReport(confirmDel)} onCancel={() => setConfirmDel(null)} />
      )}
      {confirmDelAnalysis && (
        <ConfirmDialog message="Delete this analysis? This cannot be undone."
          onConfirm={() => handleDeleteAnalysis(confirmDelAnalysis)} onCancel={() => setConfirmDelAnalysis(null)} />
      )}
    </div>
  );
}

// ─── LOGIN PAGE ─────────────────────────────────────────────────────────────────
function LoginPage({ onLogin, onSwitchToRegister }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [passcode, setPasscode] = useState('');
  const [usePasscode, setUsePasscode] = useState(false);
  const [remember, setRemember] = useState(false);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleLogin = async (e) => {
    e.preventDefault();
    setError('');
    setLoading(true);
    const users = getUsers();
    const user = users.find(u => u.username.toLowerCase() === username.toLowerCase());
    if (!user) { setError('User not found.'); setLoading(false); return; }

    if (usePasscode) {
      if (!user.passcode) { setError('No passcode set for this account. Use your password.'); setLoading(false); return; }
      const hash = await hashStr(passcode);
      if (hash !== user.passcode) { setError('Incorrect passcode.'); setLoading(false); return; }
    } else {
      const hash = await hashStr(password);
      if (hash !== user.passwordHash) { setError('Incorrect password.'); setLoading(false); return; }
    }

    setCurrentUser({ username: user.username }, remember);
    setLoading(false);
    onLogin();
  };

  return (
    <div style={{ position: 'fixed', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', background: C.bg, overflow: 'auto', padding: '2rem' }}>
      <div style={{ ...S.card, maxWidth: 420, width: '100%', padding: '2.5rem' }}>
        <div style={{ textAlign: 'center', marginBottom: 24 }}>
          <img src={logoImg} alt="My Expert Review Pal" style={{ width: 72, height: 72, borderRadius: '50%', marginBottom: 12 }} />
          <h1 style={{ fontSize: '1.4rem', color: C.primary, margin: 0 }}><BionicText>My Expert Review Pal</BionicText></h1>
          <p style={{ fontSize: '0.9rem', color: '#666', marginTop: 4 }}><BionicText>Sign in to continue</BionicText></p>
        </div>

        <form onSubmit={handleLogin}>
          <label htmlFor="login-user" style={S.label}><BionicText>Username</BionicText></label>
          <input id="login-user" style={{ ...S.input, marginBottom: 12 }} value={username}
            onChange={e => setUsername(e.target.value)} placeholder="Enter username" required autoFocus />

          {!usePasscode ? (
            <>
              <label htmlFor="login-pass" style={S.label}>Password</label>
              <input id="login-pass" type="password" style={{ ...S.input, marginBottom: 12 }} value={password}
                onChange={e => setPassword(e.target.value)} placeholder="Enter password" required />
            </>
          ) : (
            <>
              <label htmlFor="login-pin" style={S.label}>Passcode (4-6 digit PIN)</label>
              <input id="login-pin" type="password" inputMode="numeric" pattern="[0-9]{4,6}" maxLength={6}
                style={{ ...S.input, marginBottom: 12, letterSpacing: 8, textAlign: 'center', fontSize: '1.2rem' }}
                value={passcode} onChange={e => setPasscode(e.target.value.replace(/\D/g, ''))} placeholder="••••" required />
            </>
          )}

          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: '0.85rem' }}>
              <input type="checkbox" checked={remember} onChange={e => setRemember(e.target.checked)} style={{ accentColor: C.primary }} />
              Remember me
            </label>
            <button type="button" onClick={() => { setUsePasscode(!usePasscode); setError(''); }}
              style={{ background: 'none', border: 'none', color: C.accent, cursor: 'pointer', fontSize: '0.85rem' }}>
              {usePasscode ? 'Use password instead' : 'Use passcode'}
            </button>
          </div>

          {error && <div style={{ color: C.error, fontSize: '0.85rem', marginBottom: 12, padding: '0.5rem', background: '#fce4ec', borderRadius: 4 }}>{error}</div>}

          <button type="submit" style={{ ...S.btn(), width: '100%', padding: '0.75rem' }} disabled={loading}>
            {loading ? 'Signing in...' : 'Sign In'}
          </button>
        </form>

        <p style={{ textAlign: 'center', marginTop: 16, fontSize: '0.85rem', color: '#666' }}>
          Don&apos;t have an account?{' '}
          <button onClick={onSwitchToRegister} style={{ background: 'none', border: 'none', color: C.primary, cursor: 'pointer', fontWeight: 600, fontSize: '0.85rem' }}>
            Register
          </button>
        </p>
      </div>
    </div>
  );
}

// ─── REGISTER PAGE ──────────────────────────────────────────────────────────────
function RegisterPage({ onRegister, onSwitchToLogin }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPass, setConfirmPass] = useState('');
  const [passcode, setPasscode] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleRegister = async (e) => {
    e.preventDefault();
    setError('');
    if (password.length < 6) { setError('Password must be at least 6 characters.'); return; }
    if (password !== confirmPass) { setError('Passwords do not match.'); return; }
    if (passcode && (passcode.length < 4 || passcode.length > 6)) { setError('Passcode must be 4-6 digits.'); return; }

    setLoading(true);
    const users = getUsers();
    if (users.find(u => u.username.toLowerCase() === username.toLowerCase())) {
      setError('Username already taken.'); setLoading(false); return;
    }

    const passwordHash = await hashStr(password);
    const passcodeHash = passcode ? await hashStr(passcode) : null;
    const newUser = { username, passwordHash, passcode: passcodeHash, createdAt: new Date().toISOString() };
    setUsers([...users, newUser]);
    setCurrentUser({ username }, true);
    setLoading(false);
    onRegister();
  };

  return (
    <div style={{ position: 'fixed', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', background: C.bg, overflow: 'auto', padding: '2rem' }}>
      <div style={{ ...S.card, maxWidth: 420, width: '100%', padding: '2.5rem' }}>
        <div style={{ textAlign: 'center', marginBottom: 24 }}>
          <img src={logoImg} alt="My Expert Review Pal" style={{ width: 72, height: 72, borderRadius: '50%', marginBottom: 12 }} />
          <h1 style={{ fontSize: '1.4rem', color: C.primary, margin: 0 }}>Create Account</h1>
        </div>

        <form onSubmit={handleRegister}>
          <label htmlFor="reg-user" style={S.label}>Username</label>
          <input id="reg-user" style={{ ...S.input, marginBottom: 12 }} value={username}
            onChange={e => setUsername(e.target.value)} placeholder="Choose a username" required autoFocus />

          <label htmlFor="reg-pass" style={S.label}>Password</label>
          <input id="reg-pass" type="password" style={{ ...S.input, marginBottom: 12 }} value={password}
            onChange={e => setPassword(e.target.value)} placeholder="At least 6 characters" required />

          <label htmlFor="reg-confirm" style={S.label}>Confirm Password</label>
          <input id="reg-confirm" type="password" style={{ ...S.input, marginBottom: 12 }} value={confirmPass}
            onChange={e => setConfirmPass(e.target.value)} placeholder="Re-enter password" required />

          <label htmlFor="reg-pin" style={S.label}>Passcode (optional, 4-6 digit PIN)</label>
          <input id="reg-pin" type="password" inputMode="numeric" maxLength={6}
            style={{ ...S.input, marginBottom: 6, letterSpacing: 8, textAlign: 'center', fontSize: '1.2rem' }}
            value={passcode} onChange={e => setPasscode(e.target.value.replace(/\D/g, ''))} placeholder="••••" />
          <p style={{ fontSize: '0.8rem', color: '#888', marginBottom: 16 }}>
            Set a quick-access PIN for faster login. You can always use your full password.
          </p>

          {error && <div style={{ color: C.error, fontSize: '0.85rem', marginBottom: 12, padding: '0.5rem', background: '#fce4ec', borderRadius: 4 }}>{error}</div>}

          <button type="submit" style={{ ...S.btn(), width: '100%', padding: '0.75rem' }} disabled={loading}>
            {loading ? 'Creating account...' : 'Create Account'}
          </button>
        </form>

        <p style={{ textAlign: 'center', marginTop: 16, fontSize: '0.85rem', color: '#666' }}>
          Already have an account?{' '}
          <button onClick={onSwitchToLogin} style={{ background: 'none', border: 'none', color: C.primary, cursor: 'pointer', fontWeight: 600, fontSize: '0.85rem' }}>
            Sign In
          </button>
        </p>
      </div>
    </div>
  );
}

// ─── MAIN APP ───────────────────────────────────────────────────────────────────
export default function App() {
  const [authed, setAuthed] = useState(() => !!getCurrentUser());
  const [authView, setAuthView] = useState('login');
  const [page, setPage] = useState(window.location.hash.slice(1) || '/dash');
  const [projectId, setProjectId] = useState(null);
  const { toasts, add: toast } = useToast();

  useEffect(() => {
    const handler = () => {
      const hash = window.location.hash.slice(1) || '/dash';
      setPage(hash);
    };
    window.addEventListener('hashchange', handler);
    return () => window.removeEventListener('hashchange', handler);
  }, []);

  const navigate = useCallback((path, prjId = null) => {
    setProjectId(prjId);
    setPage(path);
    window.location.hash = path;
  }, []);

  const handleLogout = () => {
    logoutUser();
    setAuthed(false);
    setAuthView('login');
  };

  // Auth gate
  if (!authed) {
    if (authView === 'register') {
      return <RegisterPage onRegister={() => setAuthed(true)} onSwitchToLogin={() => setAuthView('login')} />;
    }
    return <LoginPage onLogin={() => setAuthed(true)} onSwitchToRegister={() => setAuthView('register')} />;
  }

  const currentUser = getCurrentUser();
  const navItems = [
    { path: '/dash', label: 'Dashboard' },
    { path: '/review', label: 'New Review' },
    { path: '/cfg', label: 'Configure Areas' },
    { path: '/api', label: 'API Settings' },
    { path: '/reports', label: 'Reports' },
  ];

  return (
    <>
      {/* Skip to main content link */}
      <a href="#main-content" style={{
        position: 'absolute', left: '-9999px', top: 'auto', width: '1px', height: '1px',
        overflow: 'hidden', zIndex: 10001, padding: '0.75rem', background: C.primary, color: '#fff',
        textDecoration: 'none', borderRadius: '0 0 6px 0', fontFamily: 'Arial, sans-serif',
      }} onFocus={e => { e.target.style.left = '0'; e.target.style.width = 'auto'; e.target.style.height = 'auto'; }}
        onBlur={e => { e.target.style.left = '-9999px'; e.target.style.width = '1px'; e.target.style.height = '1px'; }}>
        Skip to main content
      </a>

      {/* Sidebar Navigation */}
      <nav aria-label="Main navigation" style={S.sidebar}>
        <div style={{ ...S.sidebarTitle, display: 'flex', alignItems: 'center', gap: 10 }}>
          <img src={logoImg} alt="" style={{ width: 36, height: 36, borderRadius: '50%' }} />
          <span>My Expert Review Pal</span>
        </div>
        {currentUser && (
          <div style={{ padding: '0.5rem 1.25rem', fontSize: '0.8rem', color: 'rgba(255,255,255,0.7)', borderBottom: '1px solid rgba(255,255,255,0.15)' }}>
            Signed in as <strong style={{ color: '#fff' }}>{currentUser.username}</strong>
          </div>
        )}
        <div style={{ padding: '1rem 0', flex: 1 }}>
          {navItems.map(n => {
            const NavIcon = NAV_ICON_MAP[n.path];
            return (
              <button key={n.path} onClick={() => navigate(n.path)}
                style={S.navBtn(page === n.path)}
                aria-current={page === n.path ? 'page' : undefined}>
                <span aria-hidden="true">{NavIcon && <NavIcon size={18}/>}</span>
                <span><BionicText>{n.label}</BionicText></span>
              </button>
            );
          })}
        </div>
        <div style={{ padding: '0.75rem 1.25rem', borderTop: '1px solid rgba(255,255,255,0.15)' }}>
          <button onClick={handleLogout}
            style={{ ...S.navBtn(false), padding: '0.5rem 0', fontSize: '0.85rem', color: 'rgba(255,255,255,0.8)' }}>
            Logout
          </button>
        </div>
      </nav>

      {/* Main Content */}
      <main id="main-content" style={S.main}>
        {page === '/dash' && <Dashboard navigate={navigate} toast={toast} />}
        {page === '/review' && (
          <ReviewPage projectId={projectId} navigate={navigate} toast={toast} key={projectId || 'new'} />
        )}
        {page === '/cfg' && <ConfigureAreas toast={toast} />}
        {page === '/api' && <ApiSettings toast={toast} />}
        {page === '/reports' && <ReportsPage toast={toast} />}
      </main>

      {/* Toast Notifications */}
      <div aria-live="polite" aria-atomic="false" style={{
        position: 'fixed', top: 20, right: 20, zIndex: 9999,
        display: 'flex', flexDirection: 'column', gap: 8,
      }}>
        {toasts.map(t => (
          <div key={t.id} style={S.toast(t.type)} role="status">{t.msg}</div>
        ))}
      </div>

      {/* Accessibility Widget */}
      <AccessibilityWidget />
    </>
  );
}
