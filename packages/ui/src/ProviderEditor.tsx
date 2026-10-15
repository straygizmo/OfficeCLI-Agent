import { useState } from "react";
import type { ByokProviderInput, ByokProviderType, ProviderSummary } from "@officecli-agent/protocol";
import { t } from "./i18n";

interface Props {
  initial?: ProviderSummary;
  onSave(input: ByokProviderInput): void;
  onCancel(): void;
}

const PRESETS: Array<{ label: string; value: Partial<ByokProviderInput> }> = [
  // OpenAI: gpt-5.6 and later are assumed, so the Responses API is used (not the legacy chat completions wire format).
  { label: "OpenAI", value: { type: "openai", baseUrl: "https://api.openai.com/v1", model: "gpt-6-luna", wireApi: "responses", modelProvider: "openai" } },
  { label: "Anthropic", value: { type: "anthropic", baseUrl: "https://api.anthropic.com", model: "claude-sonnet-4.5", modelProvider: "anthropic" } },
  { label: "Azure OpenAI", value: { type: "azure", baseUrl: "https://<resource>.openai.azure.com", azureApiVersion: "2024-10-21", modelProvider: "azure_openai" } },
  { label: "Ollama (local)", value: { type: "openai", baseUrl: "http://localhost:11434/v1", model: "llama3.1", wireApi: "completions", modelProvider: "ollama" } },
  { label: "LM Studio (local)", value: { type: "openai", baseUrl: "http://localhost:1234/v1", wireApi: "completions", modelProvider: "lm_studio" } },
];

/** Form for adding or editing a BYOK provider. The API key is never echoed back from the host. */
export function ProviderEditor({ initial, onSave, onCancel }: Props) {
  // The wire format is derived from the preset / endpoint (see host resolveProvider); it stays editable in config.json.
  const [form, setForm] = useState<ByokProviderInput>({
    name: initial?.name ?? "",
    type: (initial?.type as ByokProviderType | undefined) ?? "openai",
    baseUrl: initial?.baseUrl ?? "",
    model: initial?.model ?? "",
    apiKey: undefined,
  });
  const [error, setError] = useState("");

  const set = (patch: Partial<ByokProviderInput>) => setForm((f) => ({ ...f, ...patch }));

  const submit = () => {
    if (!form.name.trim()) return setError(t.nameRequired);
    if (!/^https?:\/\//i.test(form.baseUrl.trim())) return setError(t.baseUrlInvalid);
    setError("");
    onSave({ ...form, apiKey: form.apiKey === "" ? undefined : form.apiKey });
  };

  return (
    <div className="provider-editor">
      {!initial && (
        <label>
          {t.preset}
          <select
            defaultValue=""
            onChange={(e) => {
              const preset = PRESETS[Number(e.target.value)];
              if (preset) set({ ...preset.value, name: form.name || preset.label });
            }}
          >
            <option value="">{t.select}</option>
            {PRESETS.map((p, i) => (
              <option key={p.label} value={i}>
                {p.label}
              </option>
            ))}
          </select>
        </label>
      )}
      <label>
        {t.name}
        <input value={form.name} disabled={Boolean(initial)} onChange={(e) => set({ name: e.target.value })} />
      </label>
      <label>
        {t.type}
        <select value={form.type} onChange={(e) => set({ type: e.target.value as ByokProviderType })}>
          <option value="openai">{t.typeOpenai}</option>
          <option value="azure">{t.typeAzure}</option>
          <option value="anthropic">anthropic</option>
        </select>
      </label>
      <label>
        Base URL
        <input value={form.baseUrl} onChange={(e) => set({ baseUrl: e.target.value })} />
      </label>
      <label>
        {t.apiKey} {initial?.hasApiKey ? <span className="muted small">{t.apiKeySaved}</span> : null}
        <input type="password" value={form.apiKey ?? ""} placeholder={t.apiKeyPlaceholder} onChange={(e) => set({ apiKey: e.target.value })} />
      </label>
      <label>
        {t.defaultModelId}
        <input value={form.model ?? ""} onChange={(e) => set({ model: e.target.value })} />
      </label>
      {form.type === "azure" && (
        <label>
          Azure API version
          <input value={form.azureApiVersion ?? ""} onChange={(e) => set({ azureApiVersion: e.target.value })} />
        </label>
      )}
      {error && <div className="msg system">{error}</div>}
      <div className="row-end">
        <button onClick={onCancel}>{t.cancel}</button>
        <button className="primary" onClick={submit}>
          {t.save}
        </button>
      </div>
    </div>
  );
}
