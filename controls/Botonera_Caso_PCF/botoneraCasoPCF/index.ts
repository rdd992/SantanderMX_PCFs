import { IInputs, IOutputs } from "./generated/ManifestTypes";

/** ==== Constantes ==== */
const ETAPA_ENTITY_PLURAL = "xmsbs_etapas";
const ETAPA_ENTITY_SINGLE = "xmsbs_etapa";
const ACCION_ETAPA_ENTITY_SINGLE = "xmsbs_accionetapa";
const ACCION_ENTITY_SINGLE = "xmsbs_accion";

/** Bitácora (nueva) */
const BITACORA_ENTITY_SINGLE = "xmsbs_bitacora";
const BITACORA_SELECT =
  "createdon,xmsbs_fechafinreal,_xmsbs_caso_value,_xmsbs_etapa_value,_xmsbs_etapaanterior_value,xmsbs_name";

const FIELD_ETAPA_LOOKUP_KEY = "_xmsbs_etapa_value";
const ETAPA_SELECT = "xmsbs_name,xmsbs_orden,xmsbs_codigo,_xmsbs_flujo_value";
const ACCION_ETAPA_SELECT =
  "xmsbs_name,xmsbs_codigo,_xmsbs_etapa_value,_xmsbs_accion_value,xmsbs_notificacion,xmsbs_tipoproximaetapa,_xmsbs_etapasiguiente_value,xmsbs_estadosiguiente,xmsbs_orden";

/** ==== Tipos ==== */
type Guid = string;
interface WebApiListResult<T> { entities: T[]; }
interface EtapaEntity {
  xmsbs_etapaid?: Guid; xmsbs_name?: string; xmsbs_orden?: number; xmsbs_codigo?: string;
  _xmsbs_flujo_value?: Guid; xmsbs_etapaId?: Guid;
}
interface AccionRef { xmsbs_codigo?: string; xmsbs_name?: string; }
interface AccionEtapaEntity {
  xmsbs_accionetapaid?: Guid; xmsbs_name?: string; xmsbs_codigo?: string;
  _xmsbs_etapa_value?: Guid; _xmsbs_accion_value?: Guid;
  xmsbs_notificacion?: boolean; xmsbs_tipoproximaetapa?: number;
  _xmsbs_etapasiguiente_value?: Guid; xmsbs_estadosiguiente?: number; xmsbs_orden?: number;
  xmsbs_accion?: AccionRef;
}
interface BitacoraEntity {
  createdon?: string;
  xmsbs_fechafinreal?: string;
  _xmsbs_caso_value?: Guid;
  _xmsbs_etapa_value?: Guid;
  _xmsbs_etapaanterior_value?: Guid;
  xmsbs_name?: string;
}
type RecordBag = Record<string, unknown>;
interface XrmEntity { getId?: () => string; getEntityName?: () => string }
interface XrmData { entity?: XrmEntity; refresh?: (save?: boolean) => void }
interface XrmPage { data?: XrmData }
interface XrmLike { Page?: XrmPage }

/** === Nuevo namespace/función del JS externo === */
interface BotonesCasoNamespace { executeAccionesCaso?: (executionContext?: unknown, accionCodigo?: string) => void; }
interface HostWin { BotonesCaso?: BotonesCasoNamespace; Xrm?: XrmLike; }

export class botoneraCasoPCF implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private readonly SHOW_DEBUG = true;
  private readonly DEFAULT_CASE_ENTITY = "incident";

  private context!: ComponentFramework.Context<IInputs>;
  private container!: HTMLDivElement;

  // UI
  private root!: HTMLDivElement;
  private timelineEl!: HTMLDivElement;
  private actionsEl!: HTMLDivElement;
  private infoEl!: HTMLDivElement;

  // State
  private recordId: string | null = null;
  private etapaId: string | null = null;
  private flujoId: string | null = null;
  private caseEntityLogicalName: string | null = null;

  public init(
    context: ComponentFramework.Context<IInputs>,
    _notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.container = container;

    this.root = document.createElement("div");
    this.root.className = "botonera-root";

    this.infoEl = document.createElement("div");
    this.infoEl.style.fontSize = "13px";
    this.infoEl.style.color = "var(--muted)";
    this.infoEl.style.display = "none";
    this.root.appendChild(this.infoEl);

    this.timelineEl = document.createElement("div");
    this.timelineEl.className = "timeline";
    this.root.appendChild(this.timelineEl);

    this.actionsEl = document.createElement("div");
    this.actionsEl.className = "contrato-actions";
    this.root.appendChild(this.actionsEl);

    container.appendChild(this.root);
    this.debug("init()");
  }

  public async updateView(context: ComponentFramework.Context<IInputs>): Promise<void> {
    this.context = context;
    if (!this.context.mode.isVisible) return;

    if (!this.caseEntityLogicalName) {
      this.caseEntityLogicalName = this.getCurrentCaseEntityName() ?? this.DEFAULT_CASE_ENTITY;
      this.debug("caseEntityLogicalName detectado", this.caseEntityLogicalName);
    }

    this.recordId = this.getRecordId();
    if (!this.recordId) {
      this.renderInfo("Crea y guarda el Caso para ver las etapas y acciones.");
      this.clearUI();
      return;
    }

    try {
      // 1) Caso con lookup de etapa
      const incident = (await this.context.webAPI.retrieveRecord(
        this.caseEntityLogicalName,
        this.recordId,
        `?$select=${FIELD_ETAPA_LOOKUP_KEY}`
      )) as unknown as RecordBag;

      const etapaLookup = incident[FIELD_ETAPA_LOOKUP_KEY];
      this.etapaId = typeof etapaLookup === "string" ? etapaLookup.replace(/[{}]/g, "") : null;

      if (!this.etapaId) {
        this.renderInfo("El Caso aún no tiene etapa asignada (xmsbs_etapa).");
        this.clearUI();
        return;
      }

      // 2) Etapa actual
      const etapa = (await this.context.webAPI.retrieveRecord(
        ETAPA_ENTITY_SINGLE,
        this.etapaId,
        `?$select=${ETAPA_SELECT}`
      )) as unknown as EtapaEntity;

      this.flujoId = etapa._xmsbs_flujo_value ?? null;

      // 3) Etapas del flujo
      let etapasDelFlujo: EtapaEntity[] = [];
      if (this.flujoId) {
        const query = `?$select=xmsbs_name,xmsbs_orden,xmsbs_codigo&$filter=_xmsbs_flujo_value eq ${this.flujoId}&$orderby=xmsbs_orden asc`;
        const res = (await this.context.webAPI.retrieveMultipleRecords(
          ETAPA_ENTITY_SINGLE,
          query
        )) as unknown as WebApiListResult<EtapaEntity>;
        etapasDelFlujo = res.entities ?? [];
      }

      // 4) Acciones de la etapa actual (expand para leer código de la Acción)
      const accionesQuery =
        `?$select=${ACCION_ETAPA_SELECT}` +
        `&$filter=_xmsbs_etapa_value eq ${this.etapaId}` +
        `&$orderby=xmsbs_orden asc` +
        `&$expand=xmsbs_accion($select=xmsbs_codigo,xmsbs_name)`;
      const accionesRes = (await this.context.webAPI.retrieveMultipleRecords(
        ACCION_ETAPA_ENTITY_SINGLE,
        accionesQuery
      )) as unknown as WebApiListResult<AccionEtapaEntity>;
      const acciones = accionesRes.entities ?? [];

      // 5) Bitácoras del Caso -> etapas recorridas
      const bitacoraQuery =
        `?$select=${BITACORA_SELECT}` +
        `&$filter=_xmsbs_caso_value eq ${this.recordId}` +
        `&$orderby=createdon asc`;
      const bitRes = (await this.context.webAPI.retrieveMultipleRecords(
        BITACORA_ENTITY_SINGLE,
        bitacoraQuery
      )) as unknown as WebApiListResult<BitacoraEntity>;
      const bitacoras = bitRes.entities ?? [];

      const visitedSet = this.buildVisitedStagesSet(bitacoras);

      this.renderInfo("");
      this.renderTimeline(etapasDelFlujo, this.etapaId, visitedSet);
      this.renderActions(acciones);

    } catch (err) {
      this.debugError("updateView catch()", err);
      this.renderInfo("No fue posible obtener datos del Caso/Etapas. Reintenta.");
      this.clearUI();
    }
  }

  public getOutputs(): IOutputs { return {}; }
  public destroy(): void { this.debug("destroy()"); }

  /* =================== UI helpers =================== */

  private renderInfo(html: string) {
    if (!html) { this.infoEl.innerHTML = ""; this.infoEl.style.display = "none"; }
    else { this.infoEl.innerHTML = html; this.infoEl.style.display = "block"; }
  }
  private clearUI() { this.timelineEl.innerHTML = ""; this.actionsEl.innerHTML = ""; }

  /** === TIMELINE ===
   * - La barra roja llega al centro del círculo de la etapa actual (o al final si es la última).
   * - Los círculos en ROJO (completed) se basan en las etapas realmente recorridas según Bitácora.
   *   Si se saltó una etapa, su círculo queda gris (upcoming), aunque esté antes de la etapa actual.
   */
  private renderTimeline(etapas: EtapaEntity[], etapaActualId: string, visited: Set<string>) {
    this.timelineEl.innerHTML = "";
    if (!etapas.length) { this.renderInfo("Sin etapas para el flujo."); return; }

    const bar = document.createElement("div");
    bar.className = "timeline-bar";
    this.timelineEl.appendChild(bar);

    const track = document.createElement("div");
    track.className = "tl-track";
    bar.appendChild(track);

    const progress = document.createElement("div");
    progress.className = "tl-progress";
    bar.appendChild(progress);

    const stepsLayer = document.createElement("div");
    stepsLayer.className = "tl-steps";
    stepsLayer.style.setProperty("--steps", String(etapas.length));
    bar.appendChild(stepsLayer);

    const ids = etapas.map(e => (e.xmsbs_etapaid || e.xmsbs_etapaId || "").toLowerCase());
    const idx = Math.max(0, ids.indexOf(etapaActualId.toLowerCase()));
    const total = Math.max(1, etapas.length);

    // Fracción de la pista (barra roja) – ignora saltos a efectos de la barra
    let frac = 1;
    if (total > 1) {
      if (idx >= total - 1) frac = 1;
      else frac = (idx + 0.5) / total;
    }

    const trackWidth = track.getBoundingClientRect().width;
    progress.style.width = `${trackWidth * Math.max(0, Math.min(1, frac))}px`;

    // Steps: completed solo si esa etapa aparece en Bitácora (visitada)
    etapas.forEach((e) => {
      const step = document.createElement("div");
      step.className = "tl-step";

      const stepId = (e.xmsbs_etapaid || e.xmsbs_etapaId || "").toLowerCase();
      if (stepId === etapaActualId.toLowerCase()) {
        step.classList.add("active");
      } else if (visited.has(stepId)) {
        step.classList.add("completed");
      } else {
        step.classList.add("upcoming");
      }

      step.innerHTML = `
        <div class="tl-dot"></div>
        <div class="tl-label">${this.escape(e.xmsbs_name ?? "")}</div>
      `;
      stepsLayer.appendChild(step);
    });
  }

  /** Botones: invocan JS externo con el CÓDIGO de la Acción asociada */
  private renderActions(acciones: AccionEtapaEntity[]) {
    this.actionsEl.innerHTML = "";
    if (!acciones.length) { this.renderInfo("Sin acciones configuradas para la etapa actual."); return; }

    // mismo ancho para todos los botones basado en el texto más largo
    const probe = document.createElement("span");
    probe.style.visibility = "hidden";
    probe.style.position = "absolute";
    probe.style.whiteSpace = "nowrap";
    document.body.appendChild(probe);

    let maxW = 0;
    acciones.forEach(a => { probe.textContent = a.xmsbs_name ?? ""; maxW = Math.max(maxW, probe.getBoundingClientRect().width); });
    document.body.removeChild(probe);

    const minBtnWidth = Math.ceil(maxW + 28);

    acciones.forEach((a) => {
      const nombre = a.xmsbs_name ?? "";
      const accionId = a._xmsbs_accion_value ?? null;
      const accionCodigoExpand = a.xmsbs_accion?.xmsbs_codigo ?? null;

      const btn = document.createElement("button");
      btn.className = "btn btn-fluent";
      btn.style.minWidth = `${minBtnWidth}px`;
      btn.textContent = nombre;

      btn.addEventListener("click", async () => {
        try {
          let codigo = accionCodigoExpand;
          if (!codigo && accionId) {
            btn.disabled = true;
            const acc = await this.context.webAPI.retrieveRecord(
              ACCION_ENTITY_SINGLE,
              accionId,
              "?$select=xmsbs_codigo"
            ) as unknown as { xmsbs_codigo?: string };
            codigo = acc?.xmsbs_codigo ?? null;
            btn.disabled = false;
          }
          if (codigo && codigo.length > 0) this.invokeExternalJsWithCode(codigo);
          else await this.context.navigation.openAlertDialog({ text: "No se encontró el código de la Acción asociada." });
        } catch (e) {
          this.debugError("click acción -> obtener código de Acción", e);
          await this.context.navigation.openErrorDialog({
            message: "Error obteniendo el código de la Acción.",
            details: e instanceof Error ? e.message : String(e)
          });
        }
      });

      this.actionsEl.appendChild(btn);
    });
  }

  /* =================== Bitácora helpers =================== */
  private buildVisitedStagesSet(bitacoras: BitacoraEntity[]): Set<string> {
    const set = new Set<string>();
    bitacoras.forEach(b => {
      const e = (b._xmsbs_etapa_value ?? "").replace(/[{}]/g, "").toLowerCase();
      const ea = (b._xmsbs_etapaanterior_value ?? "").replace(/[{}]/g, "").toLowerCase();
      if (e) set.add(e);
      if (ea) set.add(ea);
    });
    return set;
  }

  /* =================== Helpers de acceso seguro =================== */
  private canAccess(win: Window): boolean {
    try {
      // Si esto no lanza, es mismo origen
      void win.location?.href;
      return true;
    } catch {
      return false;
    }
  }

  private safeGet<T>(getter: () => T): T | undefined {
    try { return getter(); } catch { return undefined; }
  }

  /* =================== JS externo (namespace/función) =================== */
  private invokeExternalJsWithCode(accionCodigo: string): void {
    try {
      const wins = this.collectCandidateWindows();
      let called = false;

      for (const w of wins) {
        const fn = this.safeGet(() => (w as unknown as HostWin).BotonesCaso?.executeAccionesCaso);
        const xrm = this.safeGet(() => (w as unknown as HostWin).Xrm);

        if (typeof fn === "function") {
          const execCtx = { getFormContext: () => (xrm?.Page ?? undefined) };
          fn(execCtx, accionCodigo);
          this.debug("invokeExternalJsWithCode -> BotonesCaso.executeAccionesCaso", accionCodigo);
          called = true;
          break;
        }
      }

      if (!called) {
        void this.context.navigation.openAlertDialog({
          text: "No se encontró 'BotonesCaso.executeAccionesCaso'. Agrega/publica la web resource xmsbs_jsBotonesCaso en el formulario."
        });
      }
    } catch (e) {
      this.debugError("invokeExternalJsWithCode catch()", e as unknown);
      void this.context.navigation.openErrorDialog({
        message: "Error ejecutando el script externo.",
        details: e instanceof Error ? e.message : String(e)
      });
    }
  }

  private collectCandidateWindows(): Window[] {
    const out: Window[] = [];
    const seen = new Set<Window>();

    const pushIf = (w: Window | null | undefined) => {
      if (!w) return;
      if (seen.has(w)) return;
      seen.add(w);
      if (this.canAccess(w)) out.push(w);
    };

    // Priorizar los más probables
    pushIf(window);
    try { if (window.parent && window.parent !== window) pushIf(window.parent as Window); } catch (e) { void 0; }

    // Explorar frames del mismo origen (sin usar window.top)
    const scanFrames = (root: Window, maxDepth = 3) => {
      const queue: { w: Window; d: number }[] = [{ w: root, d: 0 }];
      while (queue.length) {
        const { w, d } = queue.shift()!;
        if (d >= maxDepth) continue;

        if (!this.canAccess(w)) continue;
        let len = 0;
        try {
          len = (w.frames as unknown as Window & { length: number }).length ?? 0;
        } catch (e) { void 0; }

        for (let i = 0; i < len; i++) {
          try {
            const ch = (w.frames[i] as unknown) as Window;
            if (!ch) continue;
            if (this.canAccess(ch) && !seen.has(ch)) {
              seen.add(ch);
              out.push(ch);
              queue.push({ w: ch, d: d + 1 });
            }
          } catch (e) { void 0; }
        }
      }
    };

    try { scanFrames(window); } catch (e) { void 0; }
    try { if (this.canAccess(window.parent as Window)) scanFrames(window.parent as Window); } catch (e) { void 0; }

    return out;
  }

  /* =================== Core helpers =================== */

  private getCurrentCaseEntityName(): string | null {
    const ctxWithPage = this.context as unknown as { page?: { entityTypeName?: string } };
    const fromPage = ctxWithPage?.page?.entityTypeName;
    if (typeof fromPage === "string" && fromPage.length > 0) return fromPage;

    try {
      const host = (window as unknown as HostWin);
      const fromXrm = host?.Xrm?.Page?.data?.entity?.getEntityName?.();
      if (typeof fromXrm === "string" && fromXrm.length > 0) return fromXrm;
    } catch (e) { void 0; }

    return null;
  }

  private getRecordId(): string | null {
    const ctxWithPage = this.context as unknown as { page?: { entityId?: string } };
    const idFromPage = ctxWithPage?.page?.entityId;
    if (typeof idFromPage === "string" && idFromPage.length > 0) {
      return idFromPage.replace(/[{}]/g, "");
    }

    try {
      const host = (window as unknown as HostWin);
      const idFromXrm = host?.Xrm?.Page?.data?.entity?.getId?.();
      if (typeof idFromXrm === "string" && idFromXrm.length > 0) {
        return idFromXrm.replace(/[{}]/g, "");
      }
    } catch (e) { void 0; }

    return null;
  }

  private escape(s: unknown): string {
    if (s == null) return "";
    return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  }

  private debug(label: string, data?: unknown) {
    if (!this.SHOW_DEBUG) return;
    try {
      if (data !== undefined) console.log("[botoneraCasoPCF]", label, data);
      else console.log("[botoneraCasoPCF]", label);
    } catch (e) { void 0; }
  }
  private debugError(label: string, err: unknown) {
    if (!this.SHOW_DEBUG) return;
    try {
      if (err instanceof Error) {
        console.error("[botoneraCasoPCF][ERROR]", label, { name: err.name, message: err.message, stack: err.stack });
      } else {
        console.error("[botoneraCasoPCF][ERROR]", label, err);
      }
    } catch (e) { void 0; }
  }
}