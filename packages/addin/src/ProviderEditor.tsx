import { useState } from "react";
import type { ByokProviderInput, ByokProviderType, ProviderSummary } from "@officecli-agent/protocol";

interface Props {
  initial?: ProviderSummary;
  onSave(input: ByokProviderInput): void;
  onCancel(): void;
}

const PRESETS: Array<{ label: string; value: Partial<ByokProviderInput> }> = [
  { label: "OpenAI", value: { type: "openai", baseUrl: "https://api.openai.com/v1", model: "gpt-4.1", modelProvider: "openai" } },
  { label: "Anthropic", value: { type: "anthropic", baseUrl: "https://api.anthropic.com", model: "claude-sonnet-4.5", modelProvider: "anthropic" } },
  { label: "Azure OpenAI", value: { type: "azure", baseUrl: "https://<resource>.openai.azure.com", azureApiVersion: "2024-10-21", modelProvider: "azure_openai" } },
  { label: "Ollama (local)", value: { type: "openai", baseUrl: "http://localhost:11434/v1", model: "llama3.1", modelProvider: "ollama" } },
  { label: "LM Studio (local)", value: { type: "openai", baseUrl: "http://localhost:1234/v1", modelProvider: "lm_studio" } },
];

/** Form for adding or editing a BYOK provider. The API key is never echoed back from the host. */
export function ProviderEditor({ initial, onSave, onCancel }: Props) {
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
    if (!form.name.trim()) return setError("名前を入力してください。");
    if (!/^https?:\/\//i.test(form.baseUrl.trim())) return setError("Base URL は http:// または https:// で始めてください。");
    setError("");
    onSave({ ...form, apiKey: form.apiKey === "" ? undefined : form.apiKey });
  };

  return (
    <div className="provider-editor">
      {!initial && (
        <label>
          プリセット
          <select
            defaultValue=""
            onChange={(e) => {
              const preset = PRESETS[Number(e.target.value)];
              if (preset) set({ ...preset.value, name: form.name || preset.label });
            }}
          >
            <option value="">(選択)</option>
            {PRESETS.map((p, i) => (
              <option key={p.label} value={i}>
                {p.label}
              </option>
            ))}
          </select>
        </label>
      )}
      <label>
        名前
        <input value={form.name} disabled={Boolean(initial)} onChange={(e) => set({ name: e.target.value })} />
      </label>
      <label>
        種類
        <select value={form.type} onChange={(e) => set({ type: e.target.value as ByokProviderType })}>
          <option value="openai">openai (OpenAI 互換: OpenAI / Ollama / vLLM / LiteLLM …)</option>
          <option value="azure">azure (Azure OpenAI ネイティブ)</option>
          <option value="anthropic">anthropic</option>
        </select>
      </label>
      <label>
        Base URL
        <input value={form.baseUrl} onChange={(e) => set({ baseUrl: e.target.value })} />
      </label>
      <label>
        API キー {initial?.hasApiKey ? <span className="muted small">(保存済み — 空のままで維持)</span> : null}
        <input type="password" value={form.apiKey ?? ""} placeholder="ローカルモデルは空でも可" onChange={(e) => set({ apiKey: e.target.value })} />
      </label>
      <label>
        既定モデル ID
        <input value={form.model ?? ""} onChange={(e) => set({ model: e.target.value })} />
      </label>
      {form.type !== "anthropic" && (
        <label>
          Wire API
          <select value={form.wireApi ?? "completions"} onChange={(e) => set({ wireApi: e.target.value as "completions" | "responses" })}>
            <option value="completions">completions</option>
            <option value="responses">responses</option>
          </select>
        </label>
      )}
      {form.type === "azure" && (
        <label>
          Azure API version
          <input value={form.azureApiVersion ?? ""} onChange={(e) => set({ azureApiVersion: e.target.value })} />
        </label>
      )}
      {error && <div className="msg system">{error}</div>}
      <div className="row-end">
        <button onClick={onCancel}>キャンセル</button>
        <button className="primary" onClick={submit}>
          保存
        </button>
      </div>
    </div>
  );
}
