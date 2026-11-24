import { useEffect, useRef, useState } from "react";
import type { ImageModelsQuery, ImageProviderInput, ImageProviderSummary, ImageProviderType } from "@officecli-agent/protocol";
import { t } from "./i18n";

interface Props {
  initial?: ImageProviderSummary;
  /** Ask the host which image models the provider in the form offers. */
  listModels(provider: ImageModelsQuery): Promise<string[]>;
  onSave(input: ImageProviderInput): void;
  onCancel(): void;
}

const PRESETS: Array<{ label: string; value: Partial<ImageProviderInput> }> = [
  { label: "OpenAI", value: { type: "openai", baseUrl: "https://api.openai.com/v1", model: "gpt-image-2.6-flare", apiKeyEnv: "OPENAI_API_KEY", resourceNameEnv: "" } },
  // Nano Banana 2.
  { label: "Google Gemini (Nano Banana 2)", value: { type: "gemini", baseUrl: "https://generativelanguage.googleapis.com/v1beta", model: "gemini-3.1-flash-image", apiKeyEnv: "GEMINI_API_KEY", resourceNameEnv: "" } },
  // Azure: the base URL may stay blank (built from the resource name variable); the model is the deployment name.
  { label: "Azure OpenAI", value: { type: "azure", baseUrl: "", model: "", azureApiVersion: "2025-04-01-preview", apiKeyEnv: "AZURE_OPENAI_API_KEY", resourceNameEnv: "AZURE_RESOURCE_NAME" } },
];

/** Value of the dropdown entry that switches to typing the ID (never a real model id). */
const CUSTOM = "__custom__";

/** Form for the image generation provider. The API key is never echoed back from the host. */
export function ImageProviderEditor({ initial, listModels, onSave, onCancel }: Props) {
  const [form, setForm] = useState<ImageProviderInput>({
    type: initial?.type ?? "openai",
    baseUrl: initial?.baseUrl ?? "",
    model: initial?.model ?? "",
    azureApiVersion: initial?.azureApiVersion,
    apiKey: undefined,
    apiKeyEnv: initial?.apiKeyEnv ?? "",
    resourceNameEnv: initial?.resourceNameEnv ?? "",
  });
  const [error, setError] = useState("");
  // Models the provider offers (empty: not fetched, failed, or Azure) and whether the ID is typed instead of picked.
  const [models, setModels] = useState<string[]>([]);
  const [typing, setTyping] = useState(false);
  const [fetching, setFetching] = useState(false);
  const [fetchError, setFetchError] = useState("");
  // Only the answer to the latest request counts (the type may have changed while one was on its way).
  const fetchSeq = useRef(0);

  const set = (patch: Partial<ImageProviderInput>) => setForm((f) => ({ ...f, ...patch }));
  const isAzure = form.type === "azure";
  const isGemini = form.type === "gemini";

  const fetchModels = (values: ImageProviderInput) => {
    const seq = ++fetchSeq.current;
    setModels([]);
    setTyping(false);
    setFetchError("");
    if (values.type === "azure" || !/^https?:\/\//i.test(values.baseUrl.trim())) {
      setFetching(false);
      return;
    }
    setFetching(true);
    listModels({ type: values.type, baseUrl: values.baseUrl, apiKey: values.apiKey || undefined, apiKeyEnv: values.apiKeyEnv, resourceNameEnv: values.resourceNameEnv })
      .then((list) => {
        if (seq !== fetchSeq.current) return;
        setModels(list);
        // Nothing chosen yet: take the provider's first (newest) image model.
        if (list.length) setForm((f) => (f.model.trim() ? f : { ...f, model: list[0] }));
      })
      .catch((err: Error) => {
        if (seq === fetchSeq.current) setFetchError(t.modelsFetchFailed(err.message));
      })
      .finally(() => {
        if (seq === fetchSeq.current) setFetching(false);
      });
  };

  // Editing the saved provider: show its models right away.
  useEffect(() => {
    if (initial) fetchModels(form);
    return () => {
      fetchSeq.current += 1;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const submit = () => {
    const baseUrl = form.baseUrl.trim();
    const viaEnv = isAzure && Boolean(form.resourceNameEnv?.trim());
    if (baseUrl ? !/^https?:\/\//i.test(baseUrl) : !viaEnv) return setError(isAzure ? t.baseUrlInvalidAzure : t.baseUrlInvalid);
    if (!form.model.trim()) return setError(t.imageModelRequired);
    setError("");
    onSave({ ...form, apiKey: form.apiKey === "" ? undefined : form.apiKey });
  };

  const apiKeyNote = initial?.apiKeySource === "stored" ? t.apiKeySaved : initial?.apiKeySource === "env" ? t.apiKeyFromEnv : null;
  const pickFromList = models.length > 0 && !typing;
  // A saved or preset ID the provider did not list stays selectable.
  const options = form.model && !models.includes(form.model) ? [form.model, ...models] : models;

  return (
    <div className="provider-editor">
      {!initial && (
        <label>
          {t.preset}
          <select
            defaultValue=""
            onChange={(e) => {
              const preset = PRESETS[Number(e.target.value)];
              if (e.target.value === "" || !preset) return;
              const next = { ...form, ...preset.value };
              setForm(next);
              fetchModels(next);
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
        {t.type}
        <select
          value={form.type}
          onChange={(e) => {
            const next = { ...form, type: e.target.value as ImageProviderType };
            setForm(next);
            fetchModels(next);
          }}
        >
          <option value="openai">{t.imageTypeOpenai}</option>
          <option value="azure">{t.typeAzure}</option>
          <option value="gemini">{t.imageTypeGemini}</option>
        </select>
      </label>
      <label>
        Base URL
        <input value={form.baseUrl} placeholder={isAzure ? "https://<resource>.openai.azure.com" : isGemini ? "https://generativelanguage.googleapis.com/v1beta" : "https://api.openai.com/v1"} onChange={(e) => set({ baseUrl: e.target.value })} />
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
        <input type="password" value={form.apiKey ?? ""} onChange={(e) => set({ apiKey: e.target.value })} />
      </label>
      <label>
        {t.apiKeyEnv}
        <input value={form.apiKeyEnv ?? ""} placeholder={isAzure ? "AZURE_OPENAI_API_KEY" : isGemini ? "GEMINI_API_KEY" : "OPENAI_API_KEY"} onChange={(e) => set({ apiKeyEnv: e.target.value })} />
        <span className="muted small">{t.apiKeyEnvHint}</span>
      </label>
      <label>
        {isAzure ? t.imageDeployment : t.imageModelId}
        {pickFromList ? (
          <select
            value={form.model}
            onChange={(e) => {
              if (e.target.value === CUSTOM) setTyping(true);
              else set({ model: e.target.value });
            }}
          >
            {options.map((id) => (
              <option key={id} value={id}>
                {id}
              </option>
            ))}
            <option value={CUSTOM}>{t.customModel}</option>
          </select>
        ) : (
          <input value={form.model} placeholder={isAzure ? undefined : isGemini ? "gemini-3.1-flash-image" : "gpt-image-2.6-flare"} onChange={(e) => set({ model: e.target.value })} />
        )}
        {fetchError && <span className="muted small">{fetchError}</span>}
      </label>
      {!isAzure && (
        <div className="row-end">
          <button disabled={fetching} onClick={() => fetchModels(form)}>
            {fetching ? t.fetchingModels : t.fetchModels}
          </button>
        </div>
      )}
      {isAzure && (
        <label>
          Azure API version
          <input value={form.azureApiVersion ?? ""} placeholder="2025-04-01-preview" onChange={(e) => set({ azureApiVersion: e.target.value })} />
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
