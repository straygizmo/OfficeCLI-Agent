import { formatSize } from "./attachments";
import { t } from "./i18n";

export interface ChipItem {
  id?: string;
  name: string;
  size: number;
  previewUrl?: string;
  thumbnail?: string;
}

const FileIcon = () => (
  <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z" />
    <path d="M14 3v5h5" />
  </svg>
);

/** Attachment chips: in the composer (with ×) and inside a sent user message (read-only). */
export function AttachmentChips({ items, onRemove }: { items: ChipItem[]; onRemove?(id: string): void }) {
  if (items.length === 0) return null;
  return (
    <div className="chips">
      {items.map((item, i) => {
        const image = item.previewUrl ?? item.thumbnail;
        return (
          <span className="chip" key={item.id ?? `${item.name}-${i}`} title={item.name}>
            {image ? <img src={image} alt="" /> : <FileIcon />}
            <span className="name">{item.name}</span>
            <span className="muted">{formatSize(item.size)}</span>
            {onRemove && item.id && (
              <button className="chip-remove" onClick={() => onRemove(item.id!)} title={t.removeAttachment} aria-label={t.removeAttachment}>
                ×
              </button>
            )}
          </span>
        );
      })}
    </div>
  );
}
