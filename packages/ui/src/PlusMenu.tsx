import { useEffect, useRef, useState } from "react";
import type { Catalog, McpServerSummary, SkillSummary } from "@officecli-agent/protocol";
import { Toggle } from "./Toggle";
import { t } from "./i18n";

interface Props {
  disabled: boolean;
  webSearch: boolean;
  onWebSearch(next: boolean): void;
  catalog: Catalog | null;
  onMcpToggle(name: string, enabled: boolean): void;
  onSkillToggle(name: string, enabled: boolean): void;
  onAddFiles(): void;
  onOpenPath(path: string): void;
}

type View = "root" | "mcp" | "skills";

const stroke = { fill: "none", stroke: "currentColor", strokeWidth: 1.7, strokeLinecap: "round" as const, strokeLinejoin: "round" as const };

const PaperclipIcon = () => (
  <svg viewBox="0 0 24 24" width="16" height="16" {...stroke} aria-hidden="true">
    <path d="M21 11.5l-8.5 8.5a5.5 5.5 0 0 1-7.8-7.8l9-9a3.5 3.5 0 0 1 5 5l-9 9a1.5 1.5 0 0 1-2.1-2.1L16 7.5" />
  </svg>
);
const GlobeIcon = () => (
  <svg viewBox="0 0 24 24" width="16" height="16" {...stroke} aria-hidden="true">
    <circle cx="12" cy="12" r="9" />
    <path d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18" />
  </svg>
);
const ServerIcon = () => (
  <svg viewBox="0 0 24 24" width="16" height="16" {...stroke} aria-hidden="true">
    <rect x="3" y="4" width="18" height="7" rx="1.5" />
    <rect x="3" y="13" width="18" height="7" rx="1.5" />
    <path d="M7 7.5h.01M7 16.5h.01" />
  </svg>
);
const SkillIcon = () => (
  <svg viewBox="0 0 24 24" width="16" height="16" {...stroke} aria-hidden="true">
    <path d="M4 7h16M4 12h16M4 17h10" />
  </svg>
);
const CheckIcon = () => (
  <svg viewBox="0 0 24 24" width="16" height="16" {...stroke} aria-hidden="true">
    <path d="M5 12.5l4.5 4.5L19 7.5" />
  </svg>
);
const ChevronRight = () => (
  <svg viewBox="0 0 24 24" width="14" height="14" {...stroke} aria-hidden="true">
    <path d="M9 6l6 6-6 6" />
  </svg>
);
const ChevronLeft = () => (
  <svg viewBox="0 0 24 24" width="14" height="14" {...stroke} aria-hidden="true">
    <path d="M15 6l-6 6 6 6" />
  </svg>
);
const OpenIcon = () => (
  <svg viewBox="0 0 24 24" width="16" height="16" {...stroke} aria-hidden="true">
    <path d="M14 4h6v6M20 4l-9 9M19 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1h5" />
  </svg>
);

/** The "+" menu next to the composer: attachments, web search, and the Copilot CLI's MCP servers and skills. */
export function PlusMenu({ disabled, webSearch, onWebSearch, catalog, onMcpToggle, onSkillToggle, onAddFiles, onOpenPath }: Props) {
  const [open, setOpen] = useState(false);
  const [view, setView] = useState<View>("root");
  const wrapRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const onMouseDown = (e: MouseEvent) => {
      if (!wrapRef.current?.contains(e.target as Node)) setOpen(false);
    };
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Escape") setOpen(false);
    };
    document.addEventListener("mousedown", onMouseDown);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("mousedown", onMouseDown);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [open]);

  const toggleOpen = () => {
    setView("root");
    setOpen((o) => !o);
  };

  return (
    <div className="plus-wrap" ref={wrapRef}>
      <button className="icon plus" onClick={toggleOpen} disabled={disabled} title={t.plusMenu} aria-label={t.plusMenu} aria-expanded={open} aria-haspopup="menu">
        <svg viewBox="0 0 24 24" width="16" height="16" {...stroke} aria-hidden="true">
          <path d="M12 5v14M5 12h14" />
        </svg>
      </button>
      {open && (
        <div className="plus-menu" role="menu">
          {view === "root" && (
            <>
              <button className="menu-row" role="menuitem" onClick={() => { setOpen(false); onAddFiles(); }}>
                <PaperclipIcon />
                <span className="grow">{t.addFiles}</span>
              </button>
              <button className="menu-row" role="menuitemcheckbox" aria-checked={webSearch} onClick={() => onWebSearch(!webSearch)}>
                <GlobeIcon />
                <span className="grow">{t.searchWeb}</span>
                {webSearch && <span className="menu-check"><CheckIcon /></span>}
              </button>
              <button className="menu-row" role="menuitem" aria-haspopup="menu" onClick={() => setView("mcp")}>
                <ServerIcon />
                <span className="grow">{t.mcpServers}</span>
                <span className="menu-sub"><ChevronRight /></span>
              </button>
              <button className="menu-row" role="menuitem" aria-haspopup="menu" onClick={() => setView("skills")}>
                <SkillIcon />
                <span className="grow">{t.skills}</span>
                <span className="menu-sub"><ChevronRight /></span>
              </button>
            </>
          )}
          {view === "mcp" && (
            <SubMenu
              title={t.mcpServers}
              onBack={() => setView("root")}
              items={catalog?.mcpServers ?? []}
              empty={t.noMcpServers}
              render={(s: McpServerSummary) => (
                <ItemRow key={s.name} name={s.name} detail={s.type} source={s.source} enabled={s.enabled} cliDisabled={s.cliDisabled} onToggle={(next) => onMcpToggle(s.name, next)} />
              )}
              openLabel={t.openCopilotSettings}
              onOpen={catalog ? () => onOpenPath(catalog.mcpConfigPath) : undefined}
            />
          )}
          {view === "skills" && (
            <SubMenu
              title={t.skills}
              onBack={() => setView("root")}
              items={catalog?.skills ?? []}
              empty={t.noSkills}
              render={(s: SkillSummary) => (
                <ItemRow key={s.name} name={s.name} detail={s.description} source={s.source} enabled={s.enabled} cliDisabled={s.cliDisabled} onToggle={(next) => onSkillToggle(s.name, next)} />
              )}
              openLabel={t.openCopilotSettings}
              onOpen={catalog ? () => onOpenPath(catalog.skillsDir) : undefined}
            />
          )}
        </div>
      )}
    </div>
  );
}

function SubMenu<T>({ title, onBack, items, empty, render, openLabel, onOpen }: { title: string; onBack(): void; items: T[]; empty: string; render(item: T): JSX.Element; openLabel: string; onOpen?: () => void }) {
  return (
    <>
      <button className="menu-row menu-back" onClick={onBack} aria-label={t.back}>
        <ChevronLeft />
        <span className="grow menu-title">{title}</span>
      </button>
      {items.length === 0 ? <p className="menu-empty muted small">{empty}</p> : items.map(render)}
      {onOpen && (
        <button className="menu-row" role="menuitem" onClick={onOpen}>
          <OpenIcon />
          <span className="grow">{openLabel}</span>
        </button>
      )}
    </>
  );
}

function ItemRow({ name, detail, source, enabled, cliDisabled, onToggle }: { name: string; detail?: string; source: string; enabled: boolean; cliDisabled: boolean; onToggle(next: boolean): void }) {
  return (
    <div className="menu-row menu-item" title={cliDisabled ? t.cliDisabledTip : detail}>
      <span className="grow menu-item-text">
        <span className="menu-item-name">
          {name} <span className="source-badge">{source}</span>
        </span>
        {detail && <span className="menu-item-detail muted small">{detail}</span>}
      </span>
      <Toggle checked={enabled} disabled={cliDisabled} onChange={onToggle} title={cliDisabled ? t.cliDisabledTip : undefined} />
    </div>
  );
}
