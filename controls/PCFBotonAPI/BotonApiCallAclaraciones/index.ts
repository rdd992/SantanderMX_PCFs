import { IInputs, IOutputs } from "./generated/ManifestTypes";

type JsonRecord = Record<string, unknown>;

export class BotonApiCallAclaraciones
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  // ===== Config =====
  private static readonly API_URL =
    "https://224b058bd2304e15a2b940182c053c.42.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/6bd633ae34f342bf86b883c913f01074/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Lmq1bHGVmh44Vdt6kUlMbCt4JfPlXzwUyAIxwjImxbU";

  // ===== PCF plumbing =====
  private _context!: ComponentFramework.Context<IInputs>;
  private _notifyOutputChanged!: () => void;

  private _container!: HTMLDivElement;
  private _wrap!: HTMLDivElement;
  private _btn!: HTMLButtonElement;
  private _status!: HTMLDivElement;

  private _currentValue = "";
  private _outValue: string | null = null;

  private _loading = false;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._context = context;
    this._notifyOutputChanged = notifyOutputChanged;
    this._container = container;

    // Wrapper (alineación derecha estilo “contrato-actions”)
    this._wrap = document.createElement("div");
    this._wrap.className = "contrato-actions";

    // Botón principal
    this._btn = document.createElement("button");
    this._btn.type = "button";
    this._btn.className = "btn btn-primary";
    this._btn.textContent = "Consultar aclaraciones";
    this._btn.addEventListener("click", () => void this.onClick());

    // Texto de estado (feedback)
    this._status = document.createElement("div");
    this._status.className = "status-text hidden";

    this._wrap.appendChild(this._btn);

    // Contenedor total
    this._container.classList.add("app-wrapper");
    this._container.appendChild(this._wrap);
    this._container.appendChild(this._status);
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;

    const raw = (context.parameters.parametrosEntrada.raw ?? "").toString();
    this._currentValue = raw;

    // Permite sobreescribir el texto del botón desde el JSON:
    // { "buttonText": "Mi botón", ... }
    const parsed = this.safeParseJson(raw);
    const btnText = this.getString(parsed, "buttonText") || "Consultar aclaraciones";
    if (this._btn.textContent !== btnText) this._btn.textContent = btnText;

    // Habilitación: puedes forzar disabled desde JSON: { "disabled": true }
    const forceDisabled = this.getBoolean(parsed, "disabled");
    this._btn.disabled = this._loading || forceDisabled;

    // Render estado si vienen last* desde JSON
    this.renderStatusFromJson(parsed);
  }

  public getOutputs(): IOutputs {
    return {
      parametrosEntrada: this._outValue ?? this._currentValue,
    };
  }


  public destroy(): void {
    // Nada crítico; el control se destruye con el contenedor
  }

  // =========================
  // Click handler
  // =========================
  private async onClick(): Promise<void> {
    if (this._loading) return;

    const inputObj = this.safeParseJson(this._currentValue) ?? {};

    // Inputs esperados desde Canvas (en parametrosEntrada JSON)
    const buc = this.nz(this.getString(inputObj, "buc"));
    const folio = this.nz(this.getString(inputObj, "folio"));
    const fechaDesde = this.nzDate(this.getUnknown(inputObj, "fechaDesde"));
    const fechaHasta = this.nzDate(this.getUnknown(inputObj, "fechaHasta"));

    const qp: Record<string, string> = {};
    if (buc) qp["buc"] = buc;
    if (folio) qp["folio"] = folio;
    if (fechaDesde) qp["fechaDesde"] = fechaDesde;
    if (fechaHasta) qp["fechaHasta"] = fechaHasta;

    const url = this.addQueryParams(BotonApiCallAclaraciones.API_URL, qp);

    this._loading = true;
    this._btn.disabled = true;
    this.setStatusLocal("Ejecutando…", "info");

    try {
      const resp = await fetch(url, {
        method: "GET",
        headers: { Accept: "application/json" },
      });

      const status = resp.status;
      const ok = resp.ok;

      const json = await this.tryReadJson(resp);
      const rawBody = this.getUnknown(json, "body") ?? json;

      let bodyText = "";
      if (typeof rawBody === "string") {
        // si viene string, intentamos parsearlo, si no, lo dejamos como string
        const parsedBody = this.safeParseJson(rawBody);
        bodyText = parsedBody ? JSON.stringify(parsedBody, null, 2) : rawBody;
      } else if (this.isRecord(rawBody)) {
        bodyText = JSON.stringify(rawBody, null, 2);
      } else {
        bodyText = JSON.stringify(rawBody ?? {}, null, 2);
      }

      // Guardamos resultado en el mismo JSON (sin perder tus inputs)
      const next: JsonRecord = {
        ...inputObj,
        lastOk: ok,
        lastStatus: status,
        lastRunAt: new Date().toISOString(),
        lastError: ok ? "" : `HTTP ${status}`,
        lastResponse: this.capString(bodyText, 50000),
      };

      this._outValue = JSON.stringify(next);
      this._notifyOutputChanged();

      this.setStatusLocal(ok ? "Listo ✅" : `Falló (HTTP ${status})`, ok ? "ok" : "bad");
    } catch (e: unknown) {
      const msg = this.errMsg(e);
      const inputObj2 = this.safeParseJson(this._currentValue) ?? {};

      const next: JsonRecord = {
        ...inputObj2,
        lastOk: false,
        lastStatus: 0,
        lastRunAt: new Date().toISOString(),
        lastError: msg,
        lastResponse: "",
      };

      this._outValue = JSON.stringify(next);
      this._notifyOutputChanged();

      this.setStatusLocal(`Error: ${msg}`, "bad");
      console.error("[PCF] Error al ejecutar GET:", e);
    } finally {
      this._loading = false;

      const parsed = this.safeParseJson(this._currentValue);
      const forceDisabled = this.getBoolean(parsed, "disabled");
      this._btn.disabled = forceDisabled;
    }
  }

  // =========================
  // UI helpers
  // =========================
  private setStatusLocal(text: string, kind: "ok" | "bad" | "info"): void {
    this._status.textContent = text;
    this._status.classList.remove("hidden", "status-ok", "status-bad", "status-info");
    this._status.classList.add(
      kind === "ok" ? "status-ok" : kind === "bad" ? "status-bad" : "status-info"
    );
  }

  private renderStatusFromJson(obj: JsonRecord | null): void {
    if (!obj) return;

    const hasAny =
      "lastOk" in obj || "lastStatus" in obj || "lastRunAt" in obj || "lastError" in obj;

    if (!hasAny) return;

    const ok = this.getBoolean(obj, "lastOk");
    const err = this.getString(obj, "lastError");
    const status = this.getNumber(obj, "lastStatus");

    if (ok) {
      this.setStatusLocal("Última ejecución: OK ✅", "ok");
    } else if (err || status) {
      this.setStatusLocal(`Última ejecución: ${err || (status ? `HTTP ${status}` : "Error")}`, "bad");
    }
  }

  // =========================
  // Fetch helpers
  // =========================
  private async tryReadJson(resp: Response): Promise<JsonRecord> {
    try {
      const j = (await resp.json()) as unknown;
      return this.isRecord(j) ? j : {};
    } catch {
      return {};
    }
  }

  // =========================
  // Safe JSON / typing helpers
  // =========================
  private safeParseJson(text: unknown): JsonRecord | null {
    if (text == null) return null;
    const s = String(text).trim();
    if (!s) return null;
    try {
      const v = JSON.parse(s) as unknown;
      return this.isRecord(v) ? v : null;
    } catch {
      return null;
    }
  }

  private isRecord(v: unknown): v is JsonRecord {
    return typeof v === "object" && v !== null && !Array.isArray(v);
  }

  private getUnknown(obj: JsonRecord | null, key: string): unknown {
    if (!obj) return undefined;
    return obj[key];
  }

  private getString(obj: JsonRecord | null, key: string): string {
    const v = this.getUnknown(obj, key);
    return typeof v === "string" ? v : "";
  }

  private getBoolean(obj: JsonRecord | null, key: string): boolean {
    const v = this.getUnknown(obj, key);
    return v === true;
  }

  private getNumber(obj: JsonRecord | null, key: string): number | null {
    const v = this.getUnknown(obj, key);
    return typeof v === "number" && Number.isFinite(v) ? v : null;
  }

  private nz(v: string): string {
    return v.trim();
  }

  /**
   * Acepta:
   * - "2026-01-05"
   * - "2026-01-05T00:00:00Z"
   * - Date
   * Retorna "YYYY-MM-DD"
   */
  private nzDate(v: unknown): string {
    if (!v) return "";
    if (v instanceof Date && !isNaN(v.getTime())) return this.toYmd(v);

    const s = String(v).trim();
    if (!s) return "";

    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

    const d = new Date(s);
    if (!isNaN(d.getTime())) return this.toYmd(d);

    return "";
  }

  private toYmd(d: Date): string {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  private addQueryParams(baseUrl: string, params: Record<string, string>): string {
    const keys = Object.keys(params).filter((k) => params[k] !== "");
    if (keys.length === 0) return baseUrl;

    const hasQ = baseUrl.includes("?");
    const sep = hasQ ? "&" : "?";
    const qs = keys
      .map((k) => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
      .join("&");
    return `${baseUrl}${sep}${qs}`;
  }

  private capString(s: string, maxLen: number): string {
    if (!s) return "";
    if (s.length <= maxLen) return s;
    return s.slice(0, maxLen) + "\n...TRUNCATED...";
  }

  private errMsg(e: unknown): string {
    if (e instanceof Error && e.message) return e.message;
    return String(e);
  }
}
