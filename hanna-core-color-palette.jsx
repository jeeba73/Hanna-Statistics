import { useState } from "react";

const palette = {
  primary: [
    {
      name: "Navy Bar",
      hex: "#1B3A6B",
      rgb: "27, 58, 107",
      usage: "Barra navigazione principale, header",
      source: "Nav bar in tutte le schermate",
    },
    {
      name: "Hanna Blue",
      hex: "#2B6CB0",
      rgb: "43, 108, 176",
      usage: "Login card, link attivi, pulsante primario (+Add Record)",
      source: "Login screen, pulsanti azione primari",
    },
    {
      name: "Royal Blue",
      hex: "#3B82F6",
      rgb: "59, 130, 246",
      usage: "Pulsante Filters, bordi attivi, elementi interattivi",
      source: "Pulsante Filters outline nelle tabelle",
    },
  ],
  accent: [
    {
      name: "Gold Amber",
      hex: "#E5A100",
      rgb: "229, 161, 0",
      usage: "Testo nav bar attivo, pulsante Export Excel, badge warning",
      source: "Testo menù navigazione, Export Excel button",
    },
    {
      name: "Action Red",
      hex: "#DC2626",
      rgb: "220, 38, 38",
      usage: "Pulsante Export PDF, alert critici",
      source: "Export PDF button in QC Record-book",
    },
    {
      name: "Success Green",
      hex: "#16A34A",
      rgb: "22, 163, 74",
      usage: "Pulsante Insert/Add, badge status OK (✓ 0d 0h 0m)",
      source: "+Insert SFG button, status badges verdi",
    },
    {
      name: "Teal Dark",
      hex: "#0D7377",
      rgb: "13, 115, 119",
      usage: "Pulsanti History, accenti secondari",
      source: "History buttons nella tabella SFG Stock",
    },
  ],
  status: [
    {
      name: "Expired Pink",
      hex: "#FECDD3",
      rgb: "254, 205, 211",
      usage: "Righe tabella per articoli scaduti/critici",
      source: "Righe evidenziate rosa in CH-SFG Stock",
    },
    {
      name: "Warning Amber",
      hex: "#FCD34D",
      rgb: "252, 211, 77",
      usage: "Badge timer con attenzione, icone warning ⚠",
      source: "Badge giallo ⏱ 0d 0h 33m, triangolo warning",
    },
    {
      name: "OK Green Badge",
      hex: "#22C55E",
      rgb: "34, 197, 94",
      usage: "Badge status completato, conferme",
      source: "Badge verdi ✓ 0d 0h 0m nel Packing Record-book",
    },
  ],
  neutral: [
    {
      name: "Background",
      hex: "#F1F5F9",
      rgb: "241, 245, 249",
      usage: "Sfondo pagina principale",
      source: "Background generale delle pagine",
    },
    {
      name: "Tile Gray",
      hex: "#CBD5E1",
      rgb: "203, 213, 225",
      usage: "Tile dashboard moduli (SFG, RM, Equipment...)",
      source: "Griglia moduli nella homepage Hanna Core",
    },
    {
      name: "Table Header",
      hex: "#475569",
      rgb: "71, 85, 105",
      usage: "Testo header tabelle, label in grassetto",
      source: "Header colonne nelle tabelle dati",
    },
    {
      name: "Body Text",
      hex: "#1E293B",
      rgb: "30, 41, 59",
      usage: "Testo corpo principale, dati tabelle",
      source: "Contenuto celle nelle tabelle record",
    },
  ],
  splash: [
    {
      name: "Cyan Wave",
      hex: "#06B6D4",
      rgb: "6, 182, 212",
      usage: "Elemento decorativo splash, accent grafico",
      source: "Onda cyan nella splash screen Hanna Core",
    },
    {
      name: "Magenta Glow",
      hex: "#A855F7",
      rgb: "168, 85, 247",
      usage: "Elemento decorativo splash (uso limitato)",
      source: "Sfumatura viola/magenta nella splash screen",
    },
  ],
};

const groups = [
  { key: "primary", label: "Blu Primari", icon: "◆" },
  { key: "accent", label: "Azioni & Accenti", icon: "●" },
  { key: "status", label: "Stati & Feedback", icon: "◐" },
  { key: "neutral", label: "Neutri & Testo", icon: "▪" },
  { key: "splash", label: "Decorativi (Splash)", icon: "✦" },
];

function ColorSwatch({ color, onSelect, isSelected }) {
  return (
    <div
      onClick={() => onSelect(color)}
      style={{
        cursor: "pointer",
        borderRadius: "12px",
        overflow: "hidden",
        background: "#fff",
        boxShadow: isSelected
          ? `0 0 0 3px ${color.hex}, 0 8px 24px rgba(0,0,0,0.15)`
          : "0 1px 4px rgba(0,0,0,0.08)",
        transition: "all 0.2s ease",
        transform: isSelected ? "scale(1.02)" : "scale(1)",
      }}
    >
      <div
        style={{
          height: "72px",
          background: color.hex,
          position: "relative",
        }}
      >
        <span
          style={{
            position: "absolute",
            bottom: "6px",
            right: "8px",
            fontSize: "11px",
            fontFamily: "'JetBrains Mono', 'SF Mono', monospace",
            color:
              luminance(color.hex) > 0.4
                ? "rgba(0,0,0,0.55)"
                : "rgba(255,255,255,0.75)",
            letterSpacing: "0.02em",
          }}
        >
          {color.hex}
        </span>
      </div>
      <div style={{ padding: "10px 12px" }}>
        <div
          style={{
            fontSize: "13px",
            fontWeight: 600,
            color: "#1E293B",
            fontFamily: "'DM Sans', sans-serif",
            marginBottom: "2px",
          }}
        >
          {color.name}
        </div>
        <div
          style={{
            fontSize: "11px",
            color: "#64748B",
            fontFamily: "'DM Sans', sans-serif",
            lineHeight: 1.3,
          }}
        >
          {color.usage.length > 50
            ? color.usage.slice(0, 50) + "…"
            : color.usage}
        </div>
      </div>
    </div>
  );
}

function luminance(hex) {
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

function DetailPanel({ color }) {
  if (!color) return null;
  return (
    <div
      style={{
        background: "#fff",
        borderRadius: "14px",
        padding: "24px",
        boxShadow: "0 2px 12px rgba(0,0,0,0.06)",
        border: "1px solid #E2E8F0",
      }}
    >
      <div style={{ display: "flex", gap: "20px", alignItems: "flex-start" }}>
        <div
          style={{
            width: "80px",
            height: "80px",
            borderRadius: "12px",
            background: color.hex,
            flexShrink: 0,
            boxShadow: `0 4px 16px ${color.hex}44`,
          }}
        />
        <div style={{ flex: 1 }}>
          <h3
            style={{
              margin: "0 0 4px 0",
              fontSize: "18px",
              fontWeight: 700,
              color: "#1E293B",
              fontFamily: "'DM Sans', sans-serif",
            }}
          >
            {color.name}
          </h3>
          <div
            style={{
              display: "flex",
              gap: "16px",
              marginBottom: "12px",
              flexWrap: "wrap",
            }}
          >
            <code
              style={{
                fontSize: "12px",
                background: "#F1F5F9",
                padding: "3px 8px",
                borderRadius: "4px",
                color: "#475569",
                fontFamily: "'JetBrains Mono', 'SF Mono', monospace",
              }}
            >
              {color.hex}
            </code>
            <code
              style={{
                fontSize: "12px",
                background: "#F1F5F9",
                padding: "3px 8px",
                borderRadius: "4px",
                color: "#475569",
                fontFamily: "'JetBrains Mono', 'SF Mono', monospace",
              }}
            >
              rgb({color.rgb})
            </code>
          </div>
          <p
            style={{
              margin: "0 0 6px 0",
              fontSize: "13px",
              color: "#334155",
              lineHeight: 1.5,
              fontFamily: "'DM Sans', sans-serif",
            }}
          >
            <strong>Utilizzo:</strong> {color.usage}
          </p>
          <p
            style={{
              margin: 0,
              fontSize: "12px",
              color: "#94A3B8",
              lineHeight: 1.4,
              fontFamily: "'DM Sans', sans-serif",
              fontStyle: "italic",
            }}
          >
            Fonte: {color.source}
          </p>
        </div>
      </div>
    </div>
  );
}

function CSSExport() {
  const [copied, setCopied] = useState(false);
  const cssVars = `:root {
  /* Hanna Core — Primary Blues */
  --hanna-navy: #1B3A6B;
  --hanna-blue: #2B6CB0;
  --hanna-royal: #3B82F6;

  /* Actions & Accents */
  --hanna-gold: #E5A100;
  --hanna-red: #DC2626;
  --hanna-green: #16A34A;
  --hanna-teal: #0D7377;

  /* Status & Feedback */
  --hanna-expired-pink: #FECDD3;
  --hanna-warning-amber: #FCD34D;
  --hanna-ok-green: #22C55E;

  /* Neutrals */
  --hanna-bg: #F1F5F9;
  --hanna-tile-gray: #CBD5E1;
  --hanna-header-text: #475569;
  --hanna-body-text: #1E293B;

  /* Decorative */
  --hanna-cyan: #06B6D4;
  --hanna-magenta: #A855F7;
}`;

  const handleCopy = () => {
    navigator.clipboard.writeText(cssVars).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };

  return (
    <div
      style={{
        background: "#0F172A",
        borderRadius: "14px",
        padding: "20px",
        position: "relative",
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: "12px",
        }}
      >
        <span
          style={{
            fontSize: "12px",
            color: "#94A3B8",
            fontFamily: "'DM Sans', sans-serif",
            fontWeight: 600,
            letterSpacing: "0.05em",
            textTransform: "uppercase",
          }}
        >
          CSS Variables — Pronte all'uso
        </span>
        <button
          onClick={handleCopy}
          style={{
            background: copied ? "#16A34A" : "#334155",
            color: "#fff",
            border: "none",
            borderRadius: "6px",
            padding: "5px 12px",
            fontSize: "12px",
            cursor: "pointer",
            fontFamily: "'DM Sans', sans-serif",
            transition: "background 0.2s",
          }}
        >
          {copied ? "✓ Copiato!" : "Copia"}
        </button>
      </div>
      <pre
        style={{
          margin: 0,
          fontSize: "12px",
          lineHeight: 1.6,
          color: "#E2E8F0",
          fontFamily: "'JetBrains Mono', 'SF Mono', monospace",
          whiteSpace: "pre-wrap",
          overflowX: "auto",
        }}
      >
        {cssVars}
      </pre>
    </div>
  );
}

export default function HannaPalette() {
  const [selected, setSelected] = useState(null);

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "linear-gradient(135deg, #F8FAFC 0%, #EEF2F7 100%)",
        fontFamily: "'DM Sans', sans-serif",
        padding: "32px 24px",
      }}
    >
      <link
        href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap"
        rel="stylesheet"
      />

      {/* Header */}
      <div style={{ maxWidth: "960px", margin: "0 auto" }}>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: "16px",
            marginBottom: "8px",
          }}
        >
          <div
            style={{
              width: "40px",
              height: "40px",
              borderRadius: "10px",
              background: "linear-gradient(135deg, #1B3A6B, #2B6CB0)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              color: "#fff",
              fontSize: "18px",
              fontWeight: 700,
            }}
          >
            H
          </div>
          <div>
            <h1
              style={{
                margin: 0,
                fontSize: "24px",
                fontWeight: 700,
                color: "#1B3A6B",
                letterSpacing: "-0.02em",
              }}
            >
              Hanna Core — Palette Colori Corporate
            </h1>
            <p
              style={{
                margin: "2px 0 0 0",
                fontSize: "13px",
                color: "#64748B",
              }}
            >
              Estratta dagli screenshot del Record Book · Riferimento per Hanna
              Statistics
            </p>
          </div>
        </div>

        <hr
          style={{
            border: "none",
            borderTop: "1px solid #E2E8F0",
            margin: "20px 0 28px 0",
          }}
        />

        {/* Color Groups */}
        {groups.map((group) => (
          <div key={group.key} style={{ marginBottom: "32px" }}>
            <h2
              style={{
                fontSize: "14px",
                fontWeight: 600,
                color: "#64748B",
                textTransform: "uppercase",
                letterSpacing: "0.06em",
                margin: "0 0 14px 0",
                display: "flex",
                alignItems: "center",
                gap: "8px",
              }}
            >
              <span style={{ fontSize: "10px" }}>{group.icon}</span>{" "}
              {group.label}
            </h2>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))",
                gap: "12px",
              }}
            >
              {palette[group.key].map((color) => (
                <ColorSwatch
                  key={color.hex}
                  color={color}
                  onSelect={setSelected}
                  isSelected={selected?.hex === color.hex}
                />
              ))}
            </div>
          </div>
        ))}

        {/* Detail Panel */}
        {selected && (
          <div style={{ marginBottom: "32px" }}>
            <h2
              style={{
                fontSize: "14px",
                fontWeight: 600,
                color: "#64748B",
                textTransform: "uppercase",
                letterSpacing: "0.06em",
                margin: "0 0 14px 0",
              }}
            >
              Dettaglio Colore Selezionato
            </h2>
            <DetailPanel color={selected} />
          </div>
        )}

        {/* CSS Export */}
        <div style={{ marginBottom: "32px" }}>
          <h2
            style={{
              fontSize: "14px",
              fontWeight: 600,
              color: "#64748B",
              textTransform: "uppercase",
              letterSpacing: "0.06em",
              margin: "0 0 14px 0",
            }}
          >
            Export CSS
          </h2>
          <CSSExport />
        </div>

        {/* Notes */}
        <div
          style={{
            background: "#FFFBEB",
            border: "1px solid #FDE68A",
            borderRadius: "12px",
            padding: "16px 20px",
            marginBottom: "24px",
          }}
        >
          <p
            style={{
              margin: 0,
              fontSize: "13px",
              color: "#92400E",
              lineHeight: 1.6,
            }}
          >
            <strong>Nota:</strong> I valori HEX sono approssimati dall'analisi
            visiva degli screenshot fotografici del monitor. Per valori esatti
            sarebbe necessario accedere direttamente al CSS di Hanna Core o
            al brand manual Hanna Instruments. I colori decorativi (cyan,
            magenta) dalla splash screen sono opzionali per Hanna Statistics.
          </p>
        </div>
      </div>
    </div>
  );
}
