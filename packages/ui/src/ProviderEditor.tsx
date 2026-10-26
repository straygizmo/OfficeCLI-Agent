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
  { label: "OpenAI", value: { type: "openai", baseUrl: "https://api.openai.com/v1", model: "gpt-6-luna", wireApi: "responses", modelProvider: "openai", apiKeyEnv: "OPENAI_API_KEY" } },
  { label: "Anthropic", value: { type: "anthropic", baseUrl: "https://api.anthropic.com", model: "claude-sonnet-4.5", modelProvider: "anthropic", apiKeyEnv: "ANTHROPIC_API_KEY" } },
  // Azure: the base URL may stay blank; the host then builds it from the resource name variable.
  { label: "Azure OpenAI", value: { type: "azure", baseUrl: "", azureApiVersion: "2024-10-21", modelProvider: "azure_openai", apiKeyEnv: "AZURE_OPENAI_API_KEY", resourceNameEnv: "AZURE_RESOURCE_NAME" } },
  { label: "Ollama (local)", value: { type: "openai", baseUrl: "http://localhost:11434/v1", model: "llama3.1", wireApi: "completions", modelProvider: "ollama", apiKeyEnv: "" } },
  { label: "LM Studio (local)", value: { type: "openai", baseUrl: "http://localhost:1234/v1", wireApi: "completions", modelProvider: "lm_studio", apiKeyEnv: "" } },
];

const AZURE_BASE_URL_PLACEHOLDER = "https://<resource>.openai.azure.com";

/** Form for adding or editing a BYOK provider. The API key is never echoed back from the host. */
export function ProviderEditor({ initial, onSave, onCancel }: Props) {
  // The wire format is derived from the preset / endpoint (see host resolveProvider); it stays editable in config.json.
  const [form, setForm] = useState<ByokProviderInput>({
    name: initial?.name ?? "",
    type: (initial?.type as ByokProviderType | undefined) ?? "openai",
    baseUrl: initial?.baseUrl ?? "",
    model: initial?.model ?? "",
    azureApiVersion: initial?.azureApiVersion,
    apiKey: undefined,
    apiKeyEnv: initial?.apiKeyEnv ?? "",
    resourceNameEnv: initial?.resourceNameEnv ?? "",
  });
  const [error, setError] = useState("");

  const set = (patch: Partial<ByokProviderInput>) => setForm((f) => ({ ...f, ...patch }));
  const isAzure = form.type === "azure";

  const submit = () => {
    if (!form.name.trim()) return setError(t.nameRequired);
    const baseUrl = form.baseUrl.trim();
    const viaEnv = isAzure && Boolean(form.resourceNameEnv?.trim());
    if (baseUrl ? !/^https?:\/\//i.test(baseUrl) : !viaEnv) return setError(isAzure ? t.baseUrlInvalidAzure : t.baseUrlInvalid);
    setError("");
    onSave({ ...form, apiKey: form.apiKey === "" ? undefined : form.apiKey });
  };

  const apiKeyNote =
    initial?.apiKeySource === "stored" ? t.apiKeySaved : initial?.apiKeySource === "env" ? t.apiKeyFromEnv : null;

  return (
    <div className="provider-editor">
      {!initial && (
        <label>
          {t.preset}
          <select
            defaultValue=""
            onChange={(e) => {
              const preset = PRESETS[Number(e.target.value)];
              if (preset) set({ resourceNameEnv: "", ...preset.value, name: form.name || preset.label });
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
        <input
          value={form.baseUrl}
          placeholder={isAzure ? AZURE_BASE_URL_PLACEHOLDER : undefined}
          onChange={(e) => set({ baseUrl: e.target.value })}
        />
      </label>
      {isAzure && (
        <label>
          {t.resourceNameEnv}
          <input value={form.resourceNameEnv ?? ""} placeholder="AZURE_RESOURCE_NAME" onChange={(e) => set({ resourceNameEnv: e.target.value })} />
          <span className="muted small">{t.resourceNameEnvHint}</span>
        </label>
      )}
      <label>
        {t.apiKey} {apiKeyNote ? <span className="muted small">{apiKeyNote}</span> : null}
        <input type="password" value={form.apiKey ?? ""} placeholder={t.apiKeyPlaceholder} onChange={(e) => set({ apiKey: e.target.value })} />
      </label>
      <label>
        {t.apiKeyEnv}
        <input value={form.apiKeyEnv ?? ""} placeholder={isAzure ? "AZURE_OPENAI_API_KEY" : "OPENAI_API_KEY"} onChange={(e) => set({ apiKeyEnv: e.target.value })} />
        <span className="muted small">{t.apiKeyEnvHint}</span>
      </label>
      <label>
        {t.defaultModelId}
        <input value={form.model ?? ""} onChange={(e) => set({ model: e.target.value })} />
      </label>
      {isAzure && (
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
