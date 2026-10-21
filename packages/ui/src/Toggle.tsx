import type { ReactNode } from "react";

interface Props {
  checked: boolean;
  onChange(next: boolean): void;
  label?: ReactNode;
  disabled?: boolean;
  title?: string;
  className?: string;
}

/** Switch-styled checkbox. The real input stays in the DOM (keyboard + screen readers); the track is drawn in CSS. */
export function Toggle({ checked, onChange, label, disabled, title, className }: Props) {
  return (
    <label className={`toggle${disabled ? " disabled" : ""}${className ? ` ${className}` : ""}`} title={title}>
      <input type="checkbox" role="switch" aria-checked={checked} checked={checked} disabled={disabled} onChange={(e) => onChange(e.target.checked)} />
      <span className="toggle-track" aria-hidden="true">
        <span className="toggle-knob" />
      </span>
      {label !== undefined && <span className="toggle-label">{label}</span>}
    </label>
  );
}
