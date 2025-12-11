import { IInputs, IOutputs } from "./generated/ManifestTypes";
import {
  provideFluentDesignSystem,
  fluentCombobox,
  fluentOption,
  fluentCheckbox,
  accentBaseColor,
  focusStrokeOuter,
  focusStrokeInner,
  SwatchRGB,
} from "@fluentui/web-components";

provideFluentDesignSystem().register(fluentCombobox(), fluentOption(), fluentCheckbox());

// === CONFIGURACIÓN DE ENDPOINTS ===
// Personas (GET)
const PERSONAS_API_URL =
  "https://224b058bd2304e15a2b940182c053c.42.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/c1dd489275bf4d9c9f7fc3927058b0d5/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=rncVCkX5MYPbuafbLiGabuvygM9I6QAS4UH8WPz2pLo";

// Productos (GET)
const PRODUCTOS_API_URL =
  "https://224b058bd2304e15a2b940182c053c.42.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/ee0e82298a764feab627da285b4f4cf0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ryIo8QLIe0tyYa4mjcATer37ieEki708JqsMIaGb2kY";

// Movimientos (GET con ?productoId=)
const MOVIMIENTOS_API_URL =
  "https://224b058bd2304e15a2b940182c053c.42.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/84065a81234e4b16899c1b4a5bcbfd5a/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=N5F-ivVp4t_VV_6tJmxdYeY-R7-3Re7fCHD8y49JJDs";

// === Mensajes de error estandarizados ===
const ERR_MSG = {
  persona:
    "Error al obtener Datos del Cliente, favor intente nuevamente o contacte al Administrador.",
  productos:
    "Error al obtener Productos del Cliente, favor intente nuevamente o contacte al Administrador.",
  movimientos:
    "Error al obtener Movimientos del Producto, favor intente nuevamente o contacte al Administrador.",
};

// === Íconos de marca para tarjetas ===
const BRAND_ICONS = {
  amex: "https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_tarjetaamex?preview=1",
  mastercard:
    "https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_tarjetamaster?preview=1",
  visa: "https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_tarjetavisa?preview=1",
  default:
    "https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_tarjetadefault?preview=1",
};

type Dinero = { monto?: number; divisa?: string };

type Pregunta1 = {
  id: string;
  name: string;
  code?: string;
  ultimaPregunta?: boolean | number | string;
  tieneMovimientos?: boolean | number | string;
  subcategoriaId?: string | null;
  // NUEVO: configuración de movimientos asociada (lookup xmsbs_confmovimiento)
  confMovId?: string | null;
};

type Pregunta2 = {
  id: string;
  name: string;
  code?: string;
  pregunta1Id?: string;
  subcategoriaId?: string | null;
  // NUEVO: mismo campo booleano xmsbs_tienemovimientos
  tieneMovimientos?: boolean | number | string;
  // NUEVO: configuración de movimientos asociada (lookup xmsbs_confmovimiento)
  confMovId?: string | null;
};

type TipoMovCfg = {
  id: string;
  name: string;
  codigo: string;
  subcategoriaId: string | null;
  ultimaPregunta: boolean;
};

type P1MovCfg = {
  id: string;
  name: string;
  codigo: string;
  subcategoriaId: string | null;
  ultimaPregunta: boolean;
};

type P2MovCfg = {
  id: string;
  name: string;
  codigo: string;
  p1MovId: string;
  subcategoriaId: string | null;
};



export class CaseEmbedded implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private context!: ComponentFramework.Context<IInputs>;
  private container!: HTMLDivElement;
  private notifyOutputChanged!: () => void;

  // ===== Bound outputs =====
  private outMiddleName: string = "";
  private outLastName: string = "";
  private outFirstName: string = "";
  private outEjecutivoTitular: string = "";
  private outAntiguedad: string = "";
  private outEmail: string = "";
  private outMobile: string = "";

  private outJsonPersona: string = "";

  private outUsuarioBancaElectronica: boolean | undefined;
  private outTenenciaProductos: boolean | undefined;

  private outGenderCode: number | undefined;
  private outMarcaDeVulnerabilidad: number | undefined;

  private outSegmento: any;
  private outSucursal: any;
  private outSucursalAlta: any;
  private customerid: any;

  // ===== Estado visual =====
  private state = {
    loading: false,
    cliente: null as any,
    clienteError: "" as string,

    productos: [] as any[],
    productosError: "" as string,

    categoria: "Tarjeta de crédito" as
      | "Tarjeta de crédito"
      | "Tarjeta de débito"
      | "Cuentas"
      | "Créditos"
      | "Inversiones"
      | "Otros",
    productoSel: null as any,

    // Preguntas 1
    preguntas: [] as Pregunta1[],
    preg1SelId: null as string | null,
    preg1SelName: null as string | null,

    // Preguntas 2
    preguntas2: [] as Pregunta2[],
    preg2SelId: null as string | null,
    preg2SelName: null as string | null,

    // Movimientos
    movLoading: false,
    movimientos: [] as any[],
    movError: "" as string,

    // NUEVO: selección de TipoMov / P1Mov / P2Mov
    movTipoSel: null as TipoMovCfg | null,
    movP1Opciones: [] as P1MovCfg[],
    movP1SelId: null as string | null,
    movP2Opciones: [] as P2MovCfg[],
    movP2SelId: null as string | null,

    modoSoloCliente: true,
    bucId: "" as string,
    crmCustomerGuid: "" as string,

    lockSoloPersona: false,

    // Habilitación del botón "Continuar Alta"
    finalizarHabilitado: false,

    // ======== Estado de tabla de movimientos (filtros, paginación, selección) ========
    movTable: {
      searchText: "",
      filtroComercio: "",
      filtroReferencia: "",
      filtroAutorizacion: "",
      filtroPan: "",
      filtroTipoCambio: "",
      fechaDesde: "",
      fechaHasta: "",
      montoMin: "",
      montoMax: "",
      filtroDuplicados: "todos" as "todos" | "con" | "sin",
      pageIndex: 0,
      pageSize: 10,
      selected: new Set<number>(),
    },
  };


  
  // =====================================================================================
  // NUEVO: configuración y matriz de códigos válidos para movimientos + debugging
  // =====================================================================================

  // NUEVO: configuración y matriz de códigos válidos para movimientos + debugging
  private movConfigDebug: any = null;

  /**
   * Matriz de códigos permitidos actualmente (puede ser la unión completa
   * o solo los códigos del tipo seleccionado).
   */
  private movCodigosMatriz: string[] = [];

  /** Unión de todos los códigos permitidos de la configuración (todos los tipos). */
  private movCodigosMatrizAll: string[] = [];

  /** Códigos agrupados por tipo de movimiento (ej: "TMV-002" -> ["2902", "691", "713"]). */
  private movCodigosPorTipo: Record<string, string[]> = {};

  /** Mapa desde código de factura -> código de tipo (ej: "2902" -> "TMV-002"). */
  private movTipoPorCodigo: Record<string, string> = {};

  /** Tipo de movimiento fijado por la selección del usuario (ej: "TMV-002") o null. */
  private movTipoFijado: string | null = null;


  // Orquestación / autosave
  private bootDone = false;
  private bootInProgress = false;
  private apiStarted = false;
  private autoSaveDone = false;
  private shouldAutoSaveAfterPersona: boolean = true;

  // arreglo de categorías actualizado
  private categorias = [
    "Tarjeta de crédito",
    "Tarjeta de débito",
    "Cuentas",
    "Créditos",
    "Inversiones",
    "Otros",
  ];

  // mapping de códigos
  private categoriaToCodigo: Record<string, string> = {
    "Tarjeta de crédito": "P0-001",
    "Tarjeta de débito": "P0-001",
    "Cuentas": "P0-002",
    "Créditos": "P0-003",
    "Inversiones": "P0-004",
    "Otros": "P0-005",
  };

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ) {
    this.context = context;
    this.container = container;
    this.notifyOutputChanged = notifyOutputChanged;

    const brand = SwatchRGB.create(179 / 255, 0, 0); // #b30000
    accentBaseColor.setValueFor(this.container, brand);
    focusStrokeOuter.setValueFor(this.container, brand);
    focusStrokeInner.setValueFor(this.container, brand);
    this.container.classList.add("app-wrapper");

    this.state.lockSoloPersona = this.hasSubcategoriaEnCaso();

    // Delegación de eventos para la tabla avanzada (paginación/selección/búsqueda)
    this.addGlobalTableDelegatedEvents();

    this.render();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.context = context;
    this.state.lockSoloPersona = this.hasSubcategoriaEnCaso();

    const bucId = context.parameters.initialNombre?.raw?.toString().trim();
    const autostart = !!context.parameters.autostart?.raw;

    if (!this.bootDone && autostart && bucId) {
      this.boot(bucId).catch(console.error);
    }
  }

  // ================= Helpers UI / CRM =================
  private showCrmAlert(message: string) {
    try {
      const nav = (window as any)?.Xrm?.Navigation;
      if (nav?.openAlertDialog) {
        nav.openAlertDialog({ text: message });
      } else {
        alert(message);
      }
    } catch {
      alert(message);
    }
  }

  // Hook de error solicitado
  private callErrorHookCliente() {
    try {
      const win: any = window as any;
      const XrmAny = win.Xrm;
      const ec =
        win.executionContext ||
        XrmAny?.Page ||
        XrmAny?.getFormContext?.() ||
        XrmAny?.Utility?.getGlobalContext?.();
      win.Caso?.general_funcionErrorAPICliente?.(ec);
    } catch (e) {
      console.warn("[PCF] No se pudo invocar Caso.general_funcionErrorAPICliente:", e);
    }
  }

  private getBrandIconUrl(rawBrand?: string): string {
    const b = (rawBrand ?? "").toString().trim().toLowerCase();
    if (!b) return BRAND_ICONS.default;
    if (b.includes("american express") || b.includes("amex")) return BRAND_ICONS.amex;
    if (b.includes("mastercard") || b.includes("master card")) return BRAND_ICONS.mastercard;
    if (b.includes("visa")) return BRAND_ICONS.visa;
    return BRAND_ICONS.default;
  }

  private isTarjeta(p: any): boolean {
    const t = (p?.tipo || "").toString().toLowerCase();
    return t.includes("tarjeta");
  }

  // ================= Arranque =================
  private async boot(bucId: string) {
  if (this.bootInProgress || this.bootDone) return;
    this.bootInProgress = true;

    try {
      // Si jsonPersona ya está COMPLETO → no hacer nada más.
      if (this.cargarDesdeCampoJsonPersona()) {
        this.bootDone = true;
        this.render();
        return;
      }

      // Caso contrario → lógica actual
      this.apiStarted = true;
      this.setLoading(true);
      this.render();

      await this.obtenerClientePorBuc(bucId);
      this.bootDone = true;
    } finally {
      this.bootInProgress = false;
      this.setLoading(false);
      this.render();
    }
  }


  // ========= usar xmsbs_jsonpersona si existe (SOLO VALIDAR, SIN ESCRIBIR) =========
  private cargarDesdeCampoJsonPersona(rawOverride?: string): boolean {
    const raw = (rawOverride ?? this.context.parameters.jsonPersona?.raw)?.toString()?.trim();
    if (!raw) return false;

    try {
      const parsed       = JSON.parse(raw);
      const clienteSec   = parsed?.clienteapi ?? parsed ?? {};
      const datosVisible = clienteSec?.datosVisible ?? {};
      const datosBase    = clienteSec?.datosBase ?? {};

      const has = (v: any) =>
        v !== null && v !== undefined && String(v).trim() !== "";

      const norm = (v: any) =>
        String(v ?? "").trim().toLowerCase();

      const equals = (a: any, b: any) =>
        norm(a) === norm(b);

      // Helper: si el valor viene como lookup serializado (JSON string), extraer el name
      const getLookupName = (rawStr: string | null): string | null => {
        if (!rawStr) return null;
        const s = rawStr.toString().trim();
        if (!s) return null;
        try {
          const parsed = JSON.parse(s);
          if (Array.isArray(parsed) && parsed.length > 0 && parsed[0]?.name) {
            return String(parsed[0].name);
          }
        } catch {
          // Si no es JSON, usamos el string tal cual
        }
        return s;
      };

      // Leemos lo que tiene hoy el formulario de CRM:
      const formFirstNameRaw   = this.getValueFromFormAttribute("xmsbs_firstname");
      const formMiddleNameRaw  = this.getValueFromFormAttribute("xmsbs_middlename");
      const formLastNameRaw    = this.getValueFromFormAttribute("xmsbs_lastname");
      const formBucRaw         = this.getValueFromFormAttribute("xmsbs_buc");
      const formSegmentoRaw    = this.getValueFromFormAttribute("xmsbs_segmento");
      const formSucursalTitRaw = this.getValueFromFormAttribute("xmsbs_sucursalTitular");
      const formCustomerId     = this.getValueFromFormAttribute("customerid");

      const formFirstName   = formFirstNameRaw;
      const formMiddleName  = formMiddleNameRaw;
      const formLastName    = formLastNameRaw;
      const formBuc         = formBucRaw;
      const formSegmento    = getLookupName(formSegmentoRaw);
      const formSucursalTit = getLookupName(formSucursalTitRaw);

      // Valores "esperados" según jsonPersona
      const jsonFirstName   = datosVisible?.nombres;
      const jsonMiddleName  = datosVisible?.apellidoPaterno;
      const jsonLastName    = datosVisible?.apellidoMaterno;
      const jsonBuc         = datosVisible?.buc;

      const jsonSegmentoNombre =
        datosVisible?.segmento?.descripcion ??
        datosVisible?.segmentoNombre;

      const jsonSucursalTitNombre =
        datosBase?.sucursalTitular?.descripcion ??
        datosBase?.sucursalTitularNombre;

      // Regla: solo consideramos "completo y en sync" si:
      // - Hay customerid en el formulario
      // - Y para cada campo que venga en el JSON, el formulario tiene el mismo valor
      const completo =
        has(formCustomerId) &&

        // Nombres
        (!has(jsonFirstName)  || equals(formFirstName,  jsonFirstName)) &&
        (!has(jsonMiddleName) || equals(formMiddleName, jsonMiddleName)) &&
        (!has(jsonLastName)   || equals(formLastName,   jsonLastName)) &&

        // BUC
        (!has(jsonBuc) || equals(formBuc, jsonBuc)) &&

        // Segmento (si viene desde JSON)
        (!has(jsonSegmentoNombre) || equals(formSegmento, jsonSegmentoNombre)) &&

        // Sucursal titular (si viene desde JSON)
        (!has(jsonSucursalTitNombre) || equals(formSucursalTit, jsonSucursalTitNombre));

      // IMPORTANTE: esta función SOLO chequea, no escribe nada en el formulario.
      // Se usa para decidir si:
      //   - podemos saltarnos la API de Persona, y
      //   - desactivar el autosave (shouldAutoSaveAfterPersona = false)
      return !!completo;
    } catch {
      return false;
    }
  }




  // ========= API Personas (GET) =========
  private async obtenerClientePorBuc(bucId: string) {
    this.showGlobalProgress("Cargando datos del cliente…");
    try {
      const url = this.addQueryParams(PERSONAS_API_URL, { bucId });
      const resp = await fetch(url, { method: "GET" });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);

      const json = await resp.json();
      const body = json?.body ?? json;
      const clienteSec = body ?? {};

      // Validación mínima
      const datosVisible = clienteSec?.datosVisible ?? {};
      const datosBase = clienteSec?.datosBase ?? {};
      const datosCRM = clienteSec?.datosCRM ?? {};
      const hasCliente =
        !!(datosCRM?.xmsbs_firstname || datosVisible?.nombres || datosCRM?.customerid?.id || datosVisible?.buc);

      if (!hasCliente) {
        throw new Error("NO_CLIENTE");
      }

      this.mapearYExponerCamposCasoDesdeClienteSec(clienteSec);
      this.setBound("outJsonPersona", JSON.stringify({ clienteapi: clienteSec }, null, 2));

      const antiguedadTxt = this.formatAntiguedad(datosBase?.antiguedadClienteBanco);
      const bancaActiva = !!datosCRM?.xmsbs_usuariobancaelectronica;

      const cliente = {
        nombre: this.nz(datosCRM?.xmsbs_firstname) || this.nz(datosVisible?.nombres),
        primerApellido: this.nz(datosCRM?.xmsbs_middlename) || this.nz(datosVisible?.apellidoPaterno),
        segundoApellido: this.nz(datosCRM?.xmsbs_lastname) || this.nz(datosVisible?.apellidoMaterno),
        curp: this.nz(datosCRM?.xmsbs_buc) || this.nz(datosVisible?.buc),
        antiguedad: antiguedadTxt,
        sucursal: this.nz(datosCRM?.xmsbs_sucursalTitular?.name) || this.nz(datosBase?.sucursalTitular?.descripcion),
        segmento: this.nz(datosCRM?.xmsbs_segmento?.name) || this.nz(datosVisible?.segmento?.descripcion),
        superMovil: bancaActiva ? "ACTIVO" : "Desactivo",
        ejecutivoTitular: this.nz(datosCRM?.xmsbs_ejecutivotitular) || this.nz(datosBase?.ejecutivoTitular),
        subsegmento: this.nz(datosCRM?.xmsbs_modeloatencion?.name),
        _bancaActivaFlag: bancaActiva
      };

      this.state = {
        ...this.state,
        cliente,
        clienteError: "",
        productos: [],
        productoSel: null,
        preguntas: [],
        preg1SelId: null,
        preg1SelName: null,
        preguntas2: [],
        preg2SelId: null,
        preg2SelName: null,
        movimientos: [],
        movLoading: false,
        movError: "",
        modoSoloCliente: true,
        bucId,
        crmCustomerGuid: this.nz(datosCRM?.customerid?.id),
        lockSoloPersona: this.hasSubcategoriaEnCaso(),
        finalizarHabilitado: false,
      };
    } catch (err) {
      console.error("[PCF] Error al obtener cliente:", err);
      this.state.cliente = null;
      this.state.clienteError = ERR_MSG.persona;

      // Evitar autosave cuando hubo error
      this.showCrmAlert(ERR_MSG.persona);
      this.callErrorHookCliente();
      this.apiStarted = false;
      this.autoSaveDone = true;
    } finally {
      this.closeGlobalProgress();
      this.setLoading(false);
      this.render();
    }
  }

  // ========= API Productos (GET al presionar botón) =========
  private async cargarProductos() {
    if (this.state.lockSoloPersona) return;

    const bucId =
      this.state.bucId || this.context.parameters.initialNombre?.raw?.toString().trim() || "";
    const customerid =
      this.state.crmCustomerGuid ||
      (Array.isArray(this.customerid) && this.customerid[0]?.id) ||
      "";

    if (!bucId || !customerid) {
      this.showCrmAlert(ERR_MSG.productos);
      return;
    }

    this.showGlobalProgress("Cargando productos…");
    this.setLoading(true);
    this.render();

    try {
      const url = this.addQueryParams(PRODUCTOS_API_URL, { bucId, customerid });
      const resp = await fetch(url, { method: "GET" });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);

      const json = await resp.json();
      const body = json?.body ?? json;

      const productos = this.unificarProductos(body);

      this.state = {
        ...this.state,
        productos,
        productosError: "",
        modoSoloCliente: false,
        productoSel: productos[0] || null,
        preguntas: [],
        preg1SelId: null,
        preg1SelName: null,
        preguntas2: [],
        preg2SelId: null,
        preg2SelName: null,
        movimientos: [],
        movLoading: false,
        movError: "",
        finalizarHabilitado: false, // siempre deshabilitado al cargar productos
      };

      await this.cargarPreguntasParaCategoria(this.state.categoria);
    } catch (e) {
      console.error("[PCF] Error al cargar productos:", e);
      this.state.productos = [];
      this.state.productosError = ERR_MSG.productos;
      this.showCrmAlert(ERR_MSG.productos);
    } finally {
      this.setLoading(false);
      this.closeGlobalProgress();
      this.render();
    }
  }

  // ===== Unificador de productos (con brandIconUrl) =====
  private unificarProductos(body: any): any[] {
    const out: any[] = [];

    // Tarjetas de crédito
    const tcs: any[] = Array.isArray(body?.tarjetaCreditos) ? body.tarjetaCreditos : [];
    for (const x of tcs) {
      const brand = x?.indicadorMarca?.descripcion || x?.tipoTarjeta?.descripcion;
      out.push({
        categoria: "Tarjeta de crédito",
        productoId: this.maskCard(x?.numeroTarjeta) || x?.numeroContrato || "Tarjeta de Crédito",
        tipo: "Tarjeta de Crédito",
        tipoPlastico: brand,
        brandIconUrl: this.getBrandIconUrl(brand),
        estatus: x?.estatusTarjeta?.descripcion,
        saldoDisponible: this.moneyToNumber(x?.saldoDisponible),
        contratoId: x?.numeroContrato,
        raw: x,
      });
    }

    // Tarjetas de débito
    const tds: any[] = Array.isArray(body?.tarjetaDebitos) ? body.tarjetaDebitos : [];
    for (const x of tds) {
      const brand = x?.indicadorMarca?.descripcion || x?.tipoTarjeta?.descripcion;
      out.push({
        categoria: "Tarjeta de débito",
        productoId: this.maskCard(x?.numeroTarjeta) || x?.numeroContrato || "Tarjeta de Débito",
        tipo: "Tarjeta de Débito",
        tipoPlastico: brand,
        brandIconUrl: this.getBrandIconUrl(brand),
        estatus: x?.estatusTarjeta?.descripcion,
        saldoDisponible: this.moneyToNumber(x?.saldoDisponible),
        contratoId: x?.numeroContrato,
        raw: x,
      });
    }

    // Cuentas
    const cfs: any[] = Array.isArray(body?.cuentaFondos) ? body.cuentaFondos : [];
    for (const x of cfs) {
      out.push({
        categoria: "Cuentas",
        productoId: x?.numeroContrato || x?.clabe || "Cuenta",
        tipo: "Cuenta",
        tipoPlastico: "",
        brandIconUrl: this.getBrandIconUrl(""),
        estatus: x?.estatusTarjeta?.descripcion || "",
        saldoDisponible: this.moneyToNumber(x?.saldoDisponible),
        contratoId: x?.numeroContrato,
        raw: x,
      });
    }

    // Créditos
    const cre: any[] = Array.isArray(body?.creditos) ? body.creditos : [];
    for (const x of cre) {
      out.push({
        categoria: "Créditos",
        productoId: x?.numeroContrato || "Crédito",
        tipo: "Crédito",
        tipoPlastico: "",
        brandIconUrl: this.getBrandIconUrl(""),
        estatus: x?.estatusCredito?.descripcion,
        saldoDisponible: this.moneyToNumber(x?.saldoDisponible),
        contratoId: x?.numeroContrato,
        raw: x,
      });
    }

    // Inversiones
    const invs: any[] = Array.isArray(body?.inversiones) ? body.inversiones : [];
    for (const x of invs) {
      out.push({
        categoria: "Inversiones",
        productoId: x?.numeroContrato || "Inversión",
        tipo: "Inversión",
        tipoPlastico: "",
        brandIconUrl: this.getBrandIconUrl(""),
        estatus: x?.estadoInversion?.descripcion,
        saldoDisponible: this.moneyToNumber(x?.saldoDisponible || x?.montoInvertido),
        contratoId: x?.numeroContrato,
        raw: x,
      });
    }

    // Seguros → Otros
    const segs: any[] = Array.isArray(body?.seguros) ? body.seguros : [];
    for (const x of segs) {
      out.push({
        categoria: "Otros",
        productoId: x?.numeroPoliza || "Seguro",
        tipo: "Seguro",
        tipoPlastico: x?.ramo?.descripcion || "",
        brandIconUrl: this.getBrandIconUrl(""),
        estatus: x?.estadoSeguro?.descripcion,
        saldoDisponible: undefined,
        contratoId: x?.numeroPoliza,
        raw: x,
      });
    }

    return out;
  }

  // ===== Mapear JSON de persona → Outputs =====
  private mapearYExponerCamposCasoDesdeClienteSec(clienteSec: any) {
    const datosVisible = clienteSec?.datosVisible ?? {};
    const datosBase = clienteSec?.datosBase ?? {};
    const datosCRM = clienteSec?.datosCRM ?? {};

    const antiguedadTxt = this.formatAntiguedad(datosBase?.antiguedadClienteBanco);

    this.setBound("outMiddleName", this.nz(datosCRM?.xmsbs_middlename) || this.nz(datosVisible?.apellidoPaterno));
    this.setBound("outLastName", this.nz(datosCRM?.xmsbs_lastname) || this.nz(datosVisible?.apellidoMaterno));
    this.setBound("outFirstName", this.nz(datosCRM?.xmsbs_firstname) || this.nz(datosVisible?.nombres));
    this.setBound("outEjecutivoTitular", this.nz(datosCRM?.xmsbs_ejecutivotitular) || this.nz(datosBase?.ejecutivoTitular));
    this.setBound("outAntiguedad", antiguedadTxt);
    this.setBound("outEmail", this.nz(datosCRM?.emailaddress) || this.nz(datosVisible?.correoElectronico));
    this.setBound("outMobile", this.nz(datosCRM?.xmsbs_mobilephone) || this.nz(datosVisible?.numeroTelefonoCelular));

    this.setBound("outUsuarioBancaElectronica", !!datosCRM?.xmsbs_usuariobancaelectronica || !!datosVisible?.usaBancaDigital);
    this.setBound("outTenenciaProductos", !!datosCRM?.xmsbs_tenenciaproductos || !!datosVisible?.tenenciaProducto);

    const genderVal = Number(datosCRM?.xmsbs_gendercode?.value ?? 0);
    this.setBound("outGenderCode", genderVal > 0 ? genderVal : undefined);

    const vulnVal = Number(datosCRM?.xmsbs_marcadevulnerabilidad?.value ?? 0);
    this.setBound("outMarcaDeVulnerabilidad", vulnVal > 0 ? vulnVal : undefined);

    this.setBound("outSegmento", this.makeLookup(datosCRM?.xmsbs_segmento));
    this.setBound("outSucursal", this.makeLookup(datosCRM?.xmsbs_sucursalTitular));
    this.setBound("outSucursalAlta", this.makeLookup(datosCRM?.xmsbs_sucursalalta));

    const cust = datosCRM?.customerid ?? clienteSec?.customerid ?? null;
    this.setBound("customerid", this.makeLookup(cust));

    this.state.bucId = this.nz(datosVisible?.buc) || this.state.bucId;
    this.state.crmCustomerGuid = this.nz(cust?.id) || this.state.crmCustomerGuid;
  }

  // ========= RENDER =========
  private render() {
    const c = this.state.cliente;

    const headerHtml = c
      ? `
      <section class="card">
        <div class="section-header">Información general</div>
        <div class="gen-grid">
          <div class="user-block">
            <img class="ico" src="https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_contact_vector_48x48?preview=1" alt="">
            <div class="user-text">
              <div class="name">${this.safe(c.nombre)} ${this.safe(c.primerApellido)} ${this.safe(c.segundoApellido)}</div>
              <div class="muted">BUC: ${this.safe(c.curp)}</div>
            </div>
          </div>
          <div class="item">
            <img class="ico" src="https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_antiguedad_vector_48x48?preview=1" alt="">
            <div class="item-text"><div class="muted">Antigüedad</div><div class="strong">${this.safe(c.antiguedad)}</div></div>
          </div>
          <div class="item">
            <img class="ico" src="https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_sucursal_vector_48x48?preview=1" alt="">
            <div class="item-text"><div class="muted">Sucursal</div><div class="strong">${this.safe(c.sucursal)}</div></div>
          </div>
          <div class="item">
            <img class="ico" src="https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_segmento_vector_48x48?preview=1" alt="">
            <div class="item-text"><div class="muted">Segmento</div><div class="strong">${this.safe(c.segmento)}</div></div>
          </div>
          <div class="item">
            <img class="ico" src="https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_canalesDigitales_vector_48x48?preview=1" alt="">
            <div class="item-text">
              <div class="muted">Canales digitales</div>
              ${
                c._bancaActivaFlag
                  ? `<div class="strong green-dot">● ${this.safe(c.superMovil)}</div>`
                  : `<div class="strong red-dot">● ${this.safe(c.superMovil)}</div>`
              }
            </div>
          </div>
          <div class="item">
            <img class="ico" src="https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_ejecutivoTitular2_vector_48x48?preview=1" alt="">
            <div class="item-text"><div class="muted">Ejecutivo Titular</div><div class="strong">${this.safe(c.ejecutivoTitular)}</div></div>
          </div>
          <div class="item">
            <img class="ico" src="https://mxmidasacldyndev.crm.dynamics.com/WebResources/xmsbs_subsegmento_vector_48x48?preview=1" alt="">
            <div class="item-text"><div class="muted">Subsegmento</div><div class="strong">${this.safe(c.subsegmento)}</div></div>
          </div>
        </div>
      </section>
      `
      : `
      <section class="card">
        <div class="section-header">Ingreso de datos</div>
        <div style="padding:12px;">Sin datos de cliente</div>
      </section>
      `;

    const cats = this.categorias
      .map(
        (cat) => `
        <button role="tab" id="categoriaProductos" class="fluent-tab ${cat === this.state.categoria ? "active" : ""}"
                aria-selected="${cat === this.state.categoria}"
                data-cat="${cat}">${cat}</button>`
      )
      .join("");

    const botonProductos = `
        <div style="display:flex; gap:12px; align-items:center; margin:4px 0 8px;">
          <button id="btnCargarProductos" class="btn btn-primary">Cargar productos</button>
          <span class="muted" style="color:#68707a;">Usa este botón para consultar la segunda API (Productos).</span>
        </div>
      `;

    let middleSection = "";
    if (!this.state.lockSoloPersona) {
      if (this.state.modoSoloCliente || (this.state.productos?.length ?? 0) === 0) {
        middleSection = `
          <section class="card">
            ${botonProductos}
          </section>`;
      } else {
        middleSection = `
          <section class="card">
            ${botonProductos}
            <div class="fluent-tabs transparent" id="tabsCats" role="tablist">${cats}</div>
            ${this.renderSeccionProductos()}
          </section>`;
      }
    }

    this.container.innerHTML = `${headerHtml}${middleSection}`;

    // Eventos botón productos
    this.container.querySelector("#btnCargarProductos")?.addEventListener("click", () => this.cargarProductos());

    // Eventos Categorías
    this.container.querySelectorAll('#tabsCats button[data-cat]').forEach((btn) => {
      btn.addEventListener("click", async () => {
        const nuevaCat = (btn as HTMLButtonElement).dataset.cat!;
        this.state.categoria = nuevaCat as any;
        this.state.productoSel = null;
        this.state.preguntas = [];
        this.state.preguntas2 = [];
        this.state.preg1SelId = this.state.preg1SelName = null;
        this.state.preg2SelId = this.state.preg2SelName = null;
        this.state.movimientos = [];
        this.state.movError = "";
        this.state.finalizarHabilitado = false; // reset
        this.state.movTable = { 
          ...this.state.movTable, 
          selected: new Set<number>(), 
          pageIndex: 0, 
          searchText: "",
          filtroComercio: "", filtroReferencia: "", filtroAutorizacion: "", filtroPan: "", filtroTipoCambio: "",
          fechaDesde: "", fechaHasta: "", montoMin: "", montoMax: "",
          filtroDuplicados: "todos"
        };
        await this.cargarPreguntasParaCategoria(nuevaCat);
        this.render();
      });
    });

    // Wire si hay productos
    if (!this.state.modoSoloCliente && !this.state.lockSoloPersona) {
      const productosFiltrados = this.getProductosPorCategoria();
      if (!this.state.productoSel && productosFiltrados.length) {
        this.state.productoSel = productosFiltrados[0];
      }

      this.container.querySelectorAll('#tabsProds button[data-prod]').forEach((btn) => {
        btn.addEventListener("click", () => {
          const pid = (btn as HTMLButtonElement).dataset.prod!;
          this.state.productoSel = this.state.productos.find((x) => x.productoId === pid) || null;
          this.state.preguntas2 = [];
          this.state.preg1SelId = this.state.preg1SelName = null;
          this.state.preg2SelId = this.state.preg2SelName = null;
          this.state.movimientos = [];
          this.state.movError = "";
          this.state.finalizarHabilitado = false; // reset
          this.state.movTable = { 
            ...this.state.movTable, 
            selected: new Set<number>(), 
            pageIndex: 0, 
            searchText: "",
            filtroComercio: "", filtroReferencia: "", filtroAutorizacion: "", filtroPan: "", filtroTipoCambio: "",
            fechaDesde: "", fechaHasta: "", montoMin: "", montoMax: "",
            filtroDuplicados: "todos"
          };
          this.render();
        });
      });

      // Pregunta 1
      const cbFluent = this.container.querySelector("#cbFluent") as any;
      if (cbFluent) {
        const nativeCb = this.container.querySelector("#cbPregunta") as HTMLSelectElement | null;
        if (nativeCb) nativeCb.hidden = true;

        cbFluent.addEventListener("change", async () => {
          const selId = this.getFluentSelectedValue(cbFluent);
          this.state.preg1SelId = selId;
          const found = this.state.preguntas.find((p) => p.id === selId);
          this.state.preg1SelName = found?.name ?? null;

          const esUltima = this.asBool(found?.ultimaPregunta);
          if (!esUltima && selId) {
            await this.cargarPreguntas2PorPregunta1(selId);
            // Requiere P2 → botón sigue deshabilitado
            this.state.finalizarHabilitado = false;
          } else {
            this.state.preguntas2 = [];
            this.state.preg2SelId = this.state.preg2SelName = null;
            // Es última → habilitar
            this.state.finalizarHabilitado = !!selId;
          }
          this.state.movimientos = [];
          this.state.movError = "";
          this.state.movTable = { ...this.state.movTable, selected: new Set<number>(), pageIndex: 0 };
          this.render();
        });
      } else {
        const cb = this.container.querySelector("#cbPregunta") as HTMLSelectElement | null;
        cb?.addEventListener("change", async () => {
          const selId = cb.value || null;
          this.state.preg1SelId = selId;
          const found = this.state.preguntas.find((p) => p.id === selId);
          this.state.preg1SelName = found?.name ?? null;
          const esUltima = this.asBool(found?.ultimaPregunta);
          if (!esUltima && selId) {
            await this.cargarPreguntas2PorPregunta1(selId);
            this.state.finalizarHabilitado = false;
          } else {
            this.state.preguntas2 = [];
            this.state.preg2SelId = this.state.preg2SelName = null;
            this.state.finalizarHabilitado = !!selId;
          }
          this.state.movimientos = [];
          this.state.movError = "";
          this.state.movTable = { ...this.state.movTable, selected: new Set<number>(), pageIndex: 0 };
          this.render();
        });
      }

      // Pregunta 2
      const cbFluent2 = this.container.querySelector("#cbFluent2") as any;
      if (cbFluent2) {
        const nativeCb2 = this.container.querySelector("#cbPregunta2") as HTMLSelectElement | null;
        if (nativeCb2) nativeCb2.hidden = true;
        cbFluent2.addEventListener("change", () => {
          const selId = this.getFluentSelectedValue(cbFluent2);
          this.state.preg2SelId = selId;
          const found = this.state.preguntas2.find((p) => p.id === selId);
          this.state.preg2SelName = found?.name ?? null;

          this.state.movimientos = [];
          this.state.movError = "";

          const requiereMov = this.asBool(found?.tieneMovimientos);
          // Si NO requiere movimientos, se puede continuar solo con Pregunta2 seleccionada
          this.state.finalizarHabilitado = !requiereMov && !!selId;

          this.state.movTable = {
            ...this.state.movTable,
            selected: new Set<number>(),
            pageIndex: 0,
          };
          this.render();
        });
      } else {
        const cb2 = this.container.querySelector("#cbPregunta2") as HTMLSelectElement | null;
        cb2?.addEventListener("change", () => {
          const selId = cb2.value || null;
          this.state.preg2SelId = selId;
          const found = this.state.preguntas2.find((p) => p.id === selId);
          this.state.preg2SelName = found?.name ?? null;

          this.state.movimientos = [];
          this.state.movError = "";

          const requiereMov = this.asBool(found?.tieneMovimientos);
          this.state.finalizarHabilitado = !requiereMov && !!selId;

          this.state.movTable = {
            ...this.state.movTable,
            selected: new Set<number>(),
            pageIndex: 0,
          };
          this.render();
        });
      }

      // 🔹 Botones “Ver movimientos” (P1 y P2)
      const btnMov1 = this.container.querySelector("#btnMovP1") as HTMLButtonElement | null;
      const btnMov2 = this.container.querySelector("#btnMov2") as HTMLButtonElement | null;

      const handlerVerMovimientos = async () => {
        try {
          await this.cargarConfigMovimientosParaSeleccion();
          await this.cargarMovimientos();
        } catch (e) {
          console.error("[PCF] Error al cargar movimientos:", e);
          this.showCrmAlert("No se pudieron cargar los movimientos. Intenta nuevamente.");
        }
      };

      btnMov1?.addEventListener("click", handlerVerMovimientos);
      btnMov2?.addEventListener("click", handlerVerMovimientos);

      // Botón Continuar Alta
      this.container
        .querySelector("#btnFinalizarAlta")
        ?.addEventListener("click", async () => {
          await this.continuarAlta();
        });
    }
  }


  // === RENDER Productos + Preguntas + Movimientos ===
  private renderSeccionProductos(): string {
    const productosFiltrados = this.getProductosPorCategoria();
    if (!this.state.productoSel && productosFiltrados.length) {
      this.state.productoSel = productosFiltrados[0];
    }

    // Tabs de productos
    const tabsProductos = productosFiltrados
      .map(
        (p) => `
          <button role="tab" id="tiposProductos"
                  class="fluent-tab ${this.state.productoSel?.productoId === p.productoId ? "active" : ""}"
                  aria-selected="${this.state.productoSel?.productoId === p.productoId}"
                  data-prod="${this.safe(p.productoId)}">
            <span>${this.safe(p.productoId)}</span>
          </button>`
      )
      .join("");

    const p = this.state.productoSel;
    const selP1 = this.state.preguntas.find((x) => x.id === this.state.preg1SelId);
    const esUltima = this.asBool(selP1?.ultimaPregunta);
    const showP2 = !esUltima && this.state.preguntas2.length > 0;

    const selP2 = this.state.preguntas2.find(pp => pp.id === this.state.preg2SelId) || null;

    // Flags de movimientos en Pregunta1 / Pregunta2
    const p1TieneMov = this.asBool(selP1?.tieneMovimientos);
    const p2TieneMov = this.asBool(selP2?.tieneMovimientos);

    // Cuándo mostrar cada botón de movimientos
    const showBtnMovP1 = !!selP1 && p1TieneMov;          // Pregunta1 seleccionada y tieneMovimientos = Sí
    const showBtnMovP2 = !!selP2 && p2TieneMov && showP2; // Pregunta2 visible, seleccionada y tieneMovimientos = Sí

    const r = p?.raw ?? {};
    const fmt = (d:any) => this.toDateStd(d) ?? this.safe(d);

    const bloqueTarjetaCredito = p && p.categoria === "Tarjeta de crédito" ? `
        <div class="field">
          <div class="field-label">Marca / Plástico</div>
          <div class="field-value flex-brand-with-icon">
            ${p.brandIconUrl ? `<img class="brand-ico-lg" src="${this.safe(p.brandIconUrl)}" alt="Marca ${this.safe(p.tipoPlastico)}" />` : ""}
            <span>${this.safe(p.tipoPlastico)}</span>
          </div>
        </div>
        <div class="field"><div class="field-label">Tipo de tarjeta</div><div class="field-value">${this.safe(r?.tipoTarjeta?.descripcion)}</div></div>
        <div class="field"><div class="field-label">Nombre del producto</div><div class="field-value">${this.safe(r?.producto?.descripcion)}</div></div>
        <div class="field"><div class="field-label">Número de tarjeta</div><div class="field-value">${this.maskCard(r?.numeroTarjeta)}</div></div>
        <div class="field"><div class="field-label">Límite de crédito</div><div class="field-value">${this.formatMoney(r?.limiteCredito?.monto)}</div></div>
        <div class="field"><div class="field-label">Saldo disponible</div><div class="field-value">${this.formatMoney(r?.saldoDisponible?.monto)}</div></div>
        <div class="field"><div class="field-label">Fecha de corte</div><div class="field-value">${fmt(r?.fechaCorte)}</div></div>
        <div class="field"><div class="field-label">Fecha de pago</div><div class="field-value">${fmt(r?.fechaPago)}</div></div>
        <div class="field"><div class="field-label">Monto a pagar mínimo</div><div class="field-value">${this.formatMoney(r?.montoPagarMinimo?.monto)}</div></div>
        <div class="field"><div class="field-label">Monto a pagar total</div><div class="field-value">${this.formatMoney(r?.montoPagarTotal?.monto)}</div></div>
        <div class="field"><div class="field-label">Estatus</div><div class="field-value">${this.safe(r?.estatusTarjeta?.descripcion)}</div></div>
      ` : "";

    const bloqueTarjetaDebito = p && p.categoria === "Tarjeta de débito" ? `
        <div class="field">
          <div class="field-label">Marca / Plástico</div>
          <div class="field-value flex-brand-with-icon">
            ${p.brandIconUrl ? `<img class="brand-ico-lg" src="${this.safe(p.brandIconUrl)}" alt="Marca ${this.safe(p.tipoPlastico)}" />` : ""}
            <span>${this.safe(p.tipoPlastico)}</span>
          </div>
        </div>
        <div class="field"><div class="field-label">Tipo de tarjeta</div><div class="field-value">${this.safe(r?.tipoTarjeta?.descripcion)}</div></div>
        <div class="field"><div class="field-label">Número de tarjeta</div><div class="field-value">${this.maskCard(r?.numeroTarjeta)}</div></div>
        <div class="field"><div class="field-label">Saldo disponible</div><div class="field-value">${this.formatMoney(r?.saldoDisponible?.monto)}</div></div>
        <div class="field"><div class="field-label">Estatus</div><div class="field-value">${this.safe(r?.estatusTarjeta?.descripcion)}</div></div>
      ` : "";

    const bloqueCuentas = p && p.categoria === "Cuentas" ? `
        <div class="field"><div class="field-label">Número de cuenta</div><div class="field-value">${this.safe(r?.clabe || r?.numeroCuenta)}</div></div>
        <div class="field"><div class="field-label">Número de tarjeta</div><div class="field-value">${this.maskCard(r?.numeroTarjeta)}</div></div>
        <div class="field"><div class="field-label">Tipo de cuenta</div><div class="field-value">${this.safe(r?.producto?.descripcion || r?.subproducto?.descripcion)}</div></div>
        <div class="field"><div class="field-label">Saldo disponible</div><div class="field-value">${this.formatMoney(r?.saldoDisponible?.monto)}</div></div>
        <div class="field"><div class="field-label">Fecha de apertura</div><div class="field-value">${fmt(r?.fechaAltaContrato)}</div></div>
        <div class="field"><div class="field-label">Estatus cuenta</div><div class="field-value">${this.safe(r?.estatusCuenta?.descripcion || r?.estatusTarjeta?.descripcion)}</div></div>
        <div class="field"><div class="field-label">Estatus tarjeta</div><div class="field-value">${this.safe(r?.estatusTarjeta?.descripcion)}</div></div>
        <div class="field"><div class="field-label">CLABE</div><div class="field-value">${this.safe(r?.clabe)}</div></div>
        <div class="field"><div class="field-label">Restricciones</div><div class="field-value">${this.safe(r?.restricciones)}</div></div>
      ` : "";

    const bloqueCreditos = p && p.categoria === "Créditos" ? `
        <div class="field"><div class="field-label">Tipo de crédito</div><div class="field-value">${this.safe(r?.producto?.descripcion || r?.subproducto?.descripcion)}</div></div>
        <div class="field"><div class="field-label">Número de contrato</div><div class="field-value">${this.safe(r?.numeroContrato)}</div></div>
        <div class="field"><div class="field-label">Saldo insoluto</div><div class="field-value">${this.formatMoney(r?.saldoInsoluto?.monto)}</div></div>
        <div class="field"><div class="field-label">Monto original</div><div class="field-value">${this.formatMoney(r?.montoOriginal?.monto)}</div></div>
        <div class="field"><div class="field-label">Fecha de contratación</div><div class="field-value">${fmt(r?.fechaAltaContrato)}</div></div>
        <div class="field"><div class="field-label">Fecha de vencimiento</div><div class="field-value">${fmt(r?.fechaVencimiento)}</div></div>
        <div class="field"><div class="field-label">Pago mensual</div><div class="field-value">${this.formatMoney(r?.montoPagoMensual?.monto)}</div></div>
        <div class="field"><div class="field-label">Tasa de interés</div><div class="field-value">${(r?.tasaInteres!=null)? (Number(r.tasaInteres)*100).toFixed(2)+'%':''}</div></div>
        <div class="field"><div class="field-label">Estatus</div><div class="field-value">${this.safe(r?.estatusCredito?.descripcion)}</div></div>
      ` : "";

    const bloqueInversiones = p && p.categoria === "Inversiones" ? `
        <div class="field"><div class="field-label">Tipo de instrumento</div><div class="field-value">${this.safe(r?.producto?.descripcion || r?.subproducto?.descripcion)}</div></div>
        <div class="field"><div class="field-label">Número de contrato / folio</div><div class="field-value">${this.safe(r?.numeroContrato)}</div></div>
        <div class="field"><div class="field-label">Monto invertido</div><div class="field-value">${this.formatMoney(r?.montoInvertido?.monto)}</div></div>
        <div class="field"><div class="field-label">Rendimiento estimado / tasa contratada</div><div class="field-value">${(r?.rendimientoEstimadoTasaContratada!=null)? (Number(r.rendimientoEstimadoTasaContratada)*100).toFixed(2)+'%':''}</div></div>
        <div class="field"><div class="field-label">Plazo</div><div class="field-value">${this.safe(r?.plazo?.monto)}</div></div>
        <div class="field"><div class="field-label">Vencimiento</div><div class="field-value">${fmt(r?.vencimiento || r?.fechaVencimiento)}</div></div>
        <div class="field"><div class="field-label">Estado actual</div><div class="field-value">${this.safe(r?.estadoInversion?.descripcion)}</div></div>
      ` : "";

    const bloqueSeguros = p && p.categoria === "Otros" ? `
        <div class="field"><div class="field-label">Número de póliza</div><div class="field-value">${this.safe(r?.numeroPoliza)}</div></div>
        <div class="field"><div class="field-label">Ramo</div><div class="field-value">${this.safe(r?.ramo?.descripcion)}</div></div>
        <div class="field"><div class="field-label">Fecha de contratación</div><div class="field-value">${fmt(r?.fechaAltaContrato)}</div></div>
        <div class="field"><div class="field-label">Fecha de vencimiento</div><div class="field-value">${fmt(r?.fechaVencimiento)}</div></div>
        <div class="field"><div class="field-label">Estado del seguro</div><div class="field-value">${this.safe(r?.estadoSeguro?.descripcion)}</div></div>
      ` : "";

    const datosProducto = p
    ? `
          <div class="section-header">Datos del producto</div>
          <div class="product-grid">

            ${bloqueTarjetaCredito}
            ${bloqueTarjetaDebito}
            ${bloqueCuentas}
            ${bloqueCreditos}
            ${bloqueInversiones}
            ${bloqueSeguros}

            <!-- Pregunta 1 -->
            <label class="question-label" id="lblMotivo">¿Qué tipo de situación reporta el cliente?</label>
            <div class="question-value" style="display:flex; gap:10px; align-items:center;">
              <fluent-combobox id="cbFluent" class="field-label" placeholder="-- Selecciona la pregunta --" aria-labelledby="lblMotivo">
                ${
                  (this.state.preguntas.length ? this.state.preguntas : [{ id: "", name: "(Sin opciones disponibles)" }])
                    .map(
                      (o) =>
                        `<fluent-option value="${this.safe(o.id)}" ${this.state.preg1SelId === o.id ? "selected" : ""}>${this.safe(o.name)}</fluent-option>`
                    )
                    .join("")
                }
              </fluent-combobox>

              <select id="cbPregunta" hidden>
                <option value="" ${!this.state.preg1SelId ? "selected" : ""}>-- Selecciona la pregunta --</option>
                ${
                  (this.state.preguntas.length ? this.state.preguntas : [{ id: "", name: "(Sin opciones disponibles)" }])
                    .map(
                      (o) =>
                        `<option value="${this.safe(o.id)}" ${this.state.preg1SelId === o.id ? "selected" : ""}>${this.safe(o.name)}</option>`
                    )
                    .join("")
                }
              </select>

              ${
                showBtnMovP1
                  ? `
                    <button
                      id="btnMovP1"
                      type="button"
                      class="btn btn-primary"
                      style="min-width:160px; white-space:nowrap; display:inline-flex; align-items:center; justify-content:center;"
                    >
                      Ver movimientos
                    </button>
                  `
                  : ""
              }
            </div>

            <!-- Pregunta 2 + Ver movimientos -->
            <label class="question-label" id="lblMotivo2" style="${showP2 ? "" : "display:none;"}">Seleccione el motivo relacionado</label>
            <div class="question-value" style="${showP2 ? "display:flex; gap:10px; align-items:center;" : "display:none;"}">
              <fluent-combobox id="cbFluent2" class="field-label" placeholder="-- Selecciona la pregunta 2 --" aria-labelledby="lblMotivo2">
                ${
                  (this.state.preguntas2.length ? this.state.preguntas2 : [{ id: "", name: "(Sin opciones disponibles)" }])
                    .map(
                      (o) =>
                        `<fluent-option value="${this.safe(o.id)}" ${this.state.preg2SelId === o.id ? "selected" : ""}>${this.safe(o.name)}</fluent-option>`
                    )
                    .join("")
                }
              </fluent-combobox>

              <select id="cbPregunta2" hidden>
                <option value="" ${!this.state.preg2SelId ? "selected" : ""}>-- Selecciona la pregunta 2 --</option>
                ${
                  (this.state.preguntas2.length ? this.state.preguntas2 : [{ id: "", name: "(Sin opciones disponibles)" }])
                    .map(
                      (o) =>
                        `<option value="${this.safe(o.id)}" ${this.state.preg2SelId === o.id ? "selected" : ""}>${this.safe(o.name)}</option>`
                    )
                    .join("")
                }
              </select>

              ${
                showBtnMovP2
                  ? `
                    <button
                      id="btnMov2"
                      type="button"
                      class="btn btn-primary"
                      style="min-width:160px; white-space:nowrap; display:inline-flex; align-items:center; justify-content:center;"
                    >
                      Ver movimientos
                    </button>
                  `
                  : ""
              }
            </div>
          </div>

          ${this.renderMovimientos()}

          <!-- SIEMPRE visible: botón Continuar Alta (habilitado según reglas) -->
          <div class="contrato-actions" style="margin-top:12px; display:flex; justify-content:flex-end;">
            <button id="btnFinalizarAlta" class="btn btn-primary" ${this.state.finalizarHabilitado ? "" : "disabled"}>
              Continuar Alta
            </button>
          </div>
        `
    : "";


    return `
        <div class="fluent-tabs transparent" id="tabsProds" role="tablist">${
          tabsProductos || "<div style='padding:8px 12px;'>Sin productos para esta categoría</div>"
        }</div>
        <div class="product-data" ${p ? "" : 'style="display:none"'} >
          ${datosProducto}
        </div>
      `;
  }


  // ========= PREGUNTAS & METADATOS =========
  private getApi() {
    return (this.context as any)?.webAPI ?? (window as any)?.Xrm?.WebApi?.online ?? (window as any)?.Xrm?.WebApi;
  }

  private async getProductoIdPorCodigo(codigo: string): Promise<string | null> {
    const api = this.getApi();
    if (!api?.retrieveMultipleRecords) return null;
    const query = `?$select=xmsbs_productoid,xmsbs_name&$filter=xmsbs_codigo eq '${codigo}'`;
    const res = await api.retrieveMultipleRecords("xmsbs_producto", query);
    const id = res?.entities?.[0]?.xmsbs_productoid as string | undefined;
    return id ?? null;
  }

  private async getProductoPorCodigo(codigo: string): Promise<{id:string,name:string,entityType:string}|null> {
    const api = this.getApi();
    if (!api?.retrieveMultipleRecords) return null;
    const query = `?$select=xmsbs_productoid,xmsbs_name&$filter=xmsbs_codigo eq '${codigo}'`;
    const res = await api.retrieveMultipleRecords("xmsbs_producto", query);
    const e = res?.entities?.[0];
    if (!e?.xmsbs_productoid) return null;
    return { id: e.xmsbs_productoid, name: e.xmsbs_name ?? "", entityType: "xmsbs_producto" };
  }

  private async getPreguntasPorProducto(productoId: string): Promise<Pregunta1[]> {
    const api = this.getApi();
    if (!api?.retrieveMultipleRecords) return [];

    const id = this.cleanGuid(productoId);
    if (!id) return [];

    const query =
      `?$select=` +
      [
        "xmsbs_pregunta1id",
        "xmsbs_name",
        "xmsbs_codigo",
        "_xmsbs_producto_value",
        "xmsbs_ultimapregunta",
        "xmsbs_tienemovimientos",
        "_xmsbs_subcategoria_value",
        // NUEVO: lookup de configuración de movimiento
        "_xmsbs_confmovimiento_value",
      ].join(",") +
      `&$filter=_xmsbs_producto_value eq ${id}` +
      `&$orderby=xmsbs_codigo asc`;

    const res = await api.retrieveMultipleRecords("xmsbs_pregunta1", query);
    return (res?.entities ?? []).map((e: any) => ({
      id: e.xmsbs_pregunta1id,
      name: e.xmsbs_name,
      code: e.xmsbs_codigo,
      ultimaPregunta: e.xmsbs_ultimapregunta,
      tieneMovimientos: e.xmsbs_tienemovimientos,
      subcategoriaId: e._xmsbs_subcategoria_value ?? null,
      // NUEVO: guardamos la conf de movimientos
      confMovId: e._xmsbs_confmovimiento_value ?? null,
    }));
  }


  private async cargarPreguntasParaCategoria(categoria: string) {
    try {
      const codigo = this.categoriaToCodigo[categoria];
      if (!codigo) {
        this.state = { ...this.state, preguntas: [], preg1SelId: null, preg1SelName: null, preguntas2: [], preg2SelId: null, preg2SelName: null, movimientos: [], movError: "", finalizarHabilitado: false };
        this.render();
        return;
      }
      const productoId = await this.getProductoIdPorCodigo(codigo);
      if (!productoId) {
        this.state = { ...this.state, preguntas: [], preg1SelId: null, preg1SelName: null, preguntas2: [], preg2SelId: null, preg2SelName: null, movimientos: [], movError: "", finalizarHabilitado: false };
        this.render();
        return;
      }
      const qs = await this.getPreguntasPorProducto(productoId);
      this.state = { ...this.state, preguntas: qs, preg1SelId: null, preg1SelName: null, preguntas2: [], preg2SelId: null, preg2SelName: null, movimientos: [], movError: "", finalizarHabilitado: false };
      this.render();
    } catch (e) {
      console.error("Error cargando preguntas para categoría:", categoria, e);
      this.state = { ...this.state, preguntas: [], preg1SelId: null, preg1SelName: null, preguntas2: [], preg2SelId: null, preg2SelName: null, movimientos: [], movError: "", finalizarHabilitado: false };
      this.render();
    }
  }

  // Preguntas 2 por Pregunta 1
  // Preguntas 2 por Pregunta 1
  private async cargarPreguntas2PorPregunta1(pregunta1Id: string) {
    const api = this.getApi();
    if (!api?.retrieveMultipleRecords) {
      this.state.preguntas2 = [];
      return;
    }

    const id = this.cleanGuid(pregunta1Id);
    if (!id) {
      this.state.preguntas2 = [];
      return;
    }

    const query =
      `?$select=` +
      [
        "xmsbs_pregunta2id",
        "xmsbs_name",
        "xmsbs_codigo",
        "_xmsbs_pregunta1_value",
        "_xmsbs_subcategoria_value",
        // NUEVO: flag tiene movimientos
        "xmsbs_tienemovimientos",
        // NUEVO: lookup configuración de movimiento
        "_xmsbs_confmovimiento_value",
      ].join(",") +
      `&$filter=_xmsbs_pregunta1_value eq ${id}` +
      `&$orderby=xmsbs_codigo asc`;

    const res = await api.retrieveMultipleRecords("xmsbs_pregunta2", query);
    const lista: Pregunta2[] = (res?.entities ?? []).map((e: any) => ({
      id: e.xmsbs_pregunta2id,
      name: e.xmsbs_name,
      code: e.xmsbs_codigo,
      pregunta1Id: e._xmsbs_pregunta1_value,
      subcategoriaId: e._xmsbs_subcategoria_value ?? null,
      // NUEVO
      tieneMovimientos: e.xmsbs_tienemovimientos,
      confMovId: e._xmsbs_confmovimiento_value ?? null,
    }));
    this.state.preguntas2 = lista;
  }

  // ========= EVENTOS delegados para tabla =========
  private addGlobalTableDelegatedEvents() {
    // Clicks (paginación, selección)
    this.container.addEventListener("click", (ev) => {
      const target = ev.target as HTMLElement;
      if (!target) return;

      // Paginación
      if (target.matches("[data-mov-page]")) {
        const action = target.getAttribute("data-mov-page")!;
        const total = this.getFilteredMovs().length;
        const { pageIndex, pageSize } = this.state.movTable;
        const lastPage = Math.max(0, Math.ceil(total / pageSize) - 1);

        if (action === "first") this.state.movTable.pageIndex = 0;
        if (action === "prev")  this.state.movTable.pageIndex = Math.max(0, pageIndex - 1);
        if (action === "next")  this.state.movTable.pageIndex = Math.min(lastPage, pageIndex + 1);
        if (action === "last")  this.state.movTable.pageIndex = lastPage;

        this.refreshMovTableUI(); // refresco parcial (no pierde foco)
      }

      // Select-all (página actual) – lo dejamos deshabilitado para evitar marcar todos
      if (target.closest?.("#mov-select-all")) {
        // No hacemos nada. El checkbox está renderizado como disabled en la UI.
        return;
      }

      // Selección de fila — ignorar duplicados y movimientos no permitidos por matriz
      if (target.closest?.(".mov-row-checkbox")) {
        const holder = target.closest(".mov-row-checkbox")!;
        const idx = Number(holder.getAttribute("data-row-index") || -1);
        if (idx >= 0) {
          const row = this.state.movimientos.find(r => r.__rowIndex === idx);
          if (!row || row.duplicado || !this.isMovimientoPermitidoPorMatriz(row)) return; // bloqueado

          const yaSeleccionado = this.state.movTable.selected.has(idx);

          if (yaSeleccionado) {
            // Quitar selección
            this.state.movTable.selected.delete(idx);

            // ✅ Si ya no queda nada seleccionado, volvemos al estado "sin restricción por matriz"
            if (this.state.movTable.selected.size === 0) {
              this.movTipoFijado = null;
              this.movCodigosMatriz = [];
            }
          } else {
            // Agregar selección
            this.state.movTable.selected.add(idx);

            // Detectar a qué tipo pertenece este movimiento (si aplica)
            const codRaw = (row.codigoFactura ?? "").toString().trim();
            const codNorm = codRaw.replace(/\D+/g, "");
            const tipo =
              this.movTipoPorCodigo[codRaw] ||
              (codNorm ? this.movTipoPorCodigo[codNorm] : undefined);

            if (tipo) {
              if (!this.movTipoFijado || this.movTipoFijado === tipo) {
                // Fijamos tipo (ej: "TMV-002") y restringimos códigos a ese tipo
                this.movTipoFijado = tipo;
                const codsTipo = this.movCodigosPorTipo[tipo] ?? [];
                if (codsTipo.length > 0) {
                  this.movCodigosMatriz = codsTipo.slice();

                  // Podar selección actual: solo dejamos seleccionados los del tipo activo
                  const allowed = new Set<string>(this.movCodigosMatriz);
                  const nuevosSel = new Set<number>();

                  for (const selIdx of this.state.movTable.selected) {
                    const rSel = this.state.movimientos.find(r => r.__rowIndex === selIdx);
                    if (!rSel) continue;
                    const rCodRaw = (rSel.codigoFactura ?? "").toString().trim();
                    const rCodNorm = rCodRaw.replace(/\D+/g, "");
                    const ok =
                      allowed.has(rCodRaw) ||
                      (rCodNorm ? allowed.has(rCodNorm) : false);
                    if (ok) nuevosSel.add(selIdx);
                  }

                  this.state.movTable.selected = nuevosSel;
                }
              } else {
                // Ya hay un tipo fijado distinto → no permitimos seleccionar este movimiento
                this.state.movTable.selected.delete(idx);
              }
            }
          }

          // Re-evaluar si se puede Continuar Alta + lógica nueva de TipoMov / P1Mov / P2Mov
          this.recomputeFinalizarPorMovimientos();
          this.refreshMovTableUI();
          (async () => { await this.onMovSelectionChanged(row); })();
        }
      }
    });

    // Inputs (búsqueda, filtros, page size, duplicados)
    this.container.addEventListener("input", (ev) => {
      const t = ev.target as HTMLInputElement | HTMLSelectElement;
      if (!t) return;

      // Búsqueda global
      if (t.id === "mov-search") {
        this.state.movTable.searchText = (t as HTMLInputElement).value ?? "";
        this.state.movTable.pageIndex = 0;
        this.refreshMovTableUI();
      }

      // Filtros de columnas / rango / duplicados
      const m = t.id.match(/^mov-filter-(.+)$/);
      if (m) {
        const key = m[1] as keyof typeof this.state.movTable;
        (this.state.movTable as any)[key] = (t as HTMLInputElement).value ?? "";
        this.state.movTable.pageIndex = 0;
        this.refreshMovTableUI();
      }

      // Tamaño de página
      if (t.id === "mov-page-size") {
        const ps = Number((t as HTMLSelectElement).value);
        if (isFinite(ps) && ps > 0) {
          this.state.movTable.pageSize = ps;
          this.state.movTable.pageIndex = 0;
          this.refreshMovTableUI();
        }
      }
    });

    // === NUEVO: cambios en combos de P1Movimiento / P2Movimiento ===
    this.container.addEventListener("change", (ev) => {
      const t = ev.target as HTMLSelectElement | null;
      if (!t) return;

      // Selección de P1Movimiento
      if (t.id === "mov-p1-select") {
        const val = t.value || "";
        this.state.movP1SelId = val || null;

        // Al cambiar P1Mov, reseteamos P2Mov
        this.state.movP2SelId = null;
        this.state.movP2Opciones = [];

        const p1Sel =
          (this.state.movP1Opciones ?? []).find(m => m.id === this.state.movP1SelId) || null;
        const esUltimaP1Mov = p1Sel ? this.asBool(p1Sel.ultimaPregunta) : false;

        (async () => {
          if (p1Sel && !esUltimaP1Mov) {
            // P1Mov no es última → cargamos P2Movimientos
            await this.cargarP2MovimientosPorP1Mov(p1Sel.id);
          }
          // Recalcular habilitación de Continuar Alta
          this.recomputeFinalizarPorMovimientos();
          this.render();
        })();
      }

      // Selección de P2Movimiento
      if (t.id === "mov-p2-select") {
        this.state.movP2SelId = t.value || null;
        this.recomputeFinalizarPorMovimientos();
        this.render();
      }
    });
  }



  // ========= MOVIMIENTOS: carga + normalización (GET) =========
  private async cargarMovimientos() {
    const prod = this.state.productoSel;
    const productoIdUI = (prod?.productoId ?? "").toString();
    const contratoIdUI = (prod?.contratoId ?? "").toString();

    if (!productoIdUI && !contratoIdUI) {
      this.showCrmAlert(ERR_MSG.movimientos);
      return;
    }

    const payloadId = /\*/.test(productoIdUI) && contratoIdUI
      ? contratoIdUI
      : (productoIdUI || contratoIdUI);

    this.showGlobalProgress("Cargando movimientos…");
    this.state.movLoading = true;
    this.state.movError = "";
    this.render();

    try {
      // Construir URL con parámetros para GET
      const urlGET = this.addQueryParams(MOVIMIENTOS_API_URL, {
        productoId: payloadId
      });

      const resp = await fetch(urlGET, {
        method: "GET",
        mode: "cors",
        cache: "no-cache",
        headers: { "Accept": "application/json" }
      });

      if (!resp.ok) {
        this.showCrmAlert(ERR_MSG.movimientos);
        throw new Error(`${ERR_MSG.movimientos} (HTTP ${resp.status})`);
      }

      const asJson = await resp.json().catch(() => ({}));
      const raw = asJson?.body ?? asJson;

      // Parsear si viene como string
      const body = typeof raw === "string" ? JSON.parse(raw) : raw;

      // IMPORTANTE: pasamos la categoría actual para filtrar tipoTransaccion
      const categoriaActual = this.state.categoria || "";
      const normalizados = this.normalizarMovimientos(body, categoriaActual);

      this.state.movimientos = normalizados;

      // Reset selección y paginación, mantener filtros
      this.state.movTable.selected = new Set<number>();
      this.state.movTable.pageIndex = 0;
      // Reset de lógica de TipoMov / P1Movimiento / P2Movimiento
      this.state.movTipoSel = null;
      this.state.movP1Opciones = [];
      this.state.movP1SelId = null;
      this.state.movP2Opciones = [];
      this.state.movP2SelId = null;
      this.recomputeFinalizarPorMovimientos();

    } catch (e: any) {
      console.error("[PCF] Error al cargar movimientos:", e);
      this.state.movimientos = [];
      // Reset de lógica de TipoMov / P1Movimiento / P2Movimiento
      this.state.movTipoSel = null;
      this.state.movP1Opciones = [];
      this.state.movP1SelId = null;
      this.state.movP2Opciones = [];
      this.state.movP2SelId = null;
      this.recomputeFinalizarPorMovimientos();
      this.state.movError = e?.message ?? ERR_MSG.movimientos;

    } finally {
      this.state.movLoading = false;
      this.closeGlobalProgress();
      this.render();
    }
  }


  /**
   * Nueva lógica: decide si el botón Continuar Alta debe estar habilitado
   * según Pregunta1/Pregunta2 y los movimientos seleccionados.
   */
  private recomputeFinalizarPorMovimientos() {
    const p1 = this.state.preguntas.find(p => p.id === this.state.preg1SelId) || null;
    const p2 = this.state.preguntas2.find(p => p.id === this.state.preg2SelId) || null;

    const tieneMovP1 = this.asBool(p1?.tieneMovimientos);
    const tieneMovP2 = this.asBool(p2?.tieneMovimientos);
    const esUltimaP1 = this.asBool(p1?.ultimaPregunta);

    const selCount = this.state.movTable.selected?.size ?? 0;
    const hayMovs = (this.state.movimientos?.length ?? 0) > 0;

    // ============================================
    // Regla especial original:
    // Si Pregunta 2 NO requiere movimientos → habilitar Continuar Alta SIN condiciones
    // ============================================
    if (p2 && !tieneMovP2) {
      this.state.finalizarHabilitado = true;
      this.syncFinalizarAltaButtonUI();
      return;
    }

    let habilitado = this.state.finalizarHabilitado;

    // Caso 1: Pregunta 2 con movimientos = Sí
    if (p2 && tieneMovP2) {
      habilitado = hayMovs && selCount > 0;
    }
    // Caso 2: Pregunta 1 NO es última y tiene movimientos = Sí
    else if (p1 && !esUltimaP1 && tieneMovP1) {
      habilitado = hayMovs && selCount > 0;
    }

    // ================================
    // NUEVO: reglas para P1Movimiento / P2Movimiento
    // Solo aplican CUANDO hay movimientos y la configuración de movs está activa.
    // ================================
    const tipoMovSel = this.state.movTipoSel;
    const p1MovList = this.state.movP1Opciones ?? [];
    const p2MovList = this.state.movP2Opciones ?? [];

    if (hayMovs && tipoMovSel) {
      // Si el tipo de movimiento NO es últimaPregunta → debo elegir al menos un P1Movimiento
      if (!this.asBool(tipoMovSel.ultimaPregunta) && p1MovList.length > 0) {
        if (!this.state.movP1SelId) {
          habilitado = false;
        } else {
          const p1MovSel = p1MovList.find(m => m.id === this.state.movP1SelId) || null;
          const esUltimaP1Mov = p1MovSel ? this.asBool(p1MovSel.ultimaPregunta) : false;

          // Si P1Movimiento tampoco es últimaPregunta y hay P2Movimiento configurado → también es obligatorio
          if (!esUltimaP1Mov && p2MovList.length > 0) {
            if (!this.state.movP2SelId) {
              habilitado = false;
            }
          }
        }
      }
    }

    this.state.finalizarHabilitado = habilitado;
    this.syncFinalizarAltaButtonUI();
  }


  // Sincroniza el estado habilitado/deshabilitado del botón "Continuar Alta"
  // con el valor actual de this.state.finalizarHabilitado sin re-render completo.
  private syncFinalizarAltaButtonUI() {
    try {
      const btn = this.container.querySelector("#btnFinalizarAlta") as HTMLButtonElement | null;
      if (!btn) return;

      if (this.state.finalizarHabilitado) {
        btn.removeAttribute("disabled");
      } else {
        btn.setAttribute("disabled", "true");
      }
    } catch (e) {
      console.warn("[PCF] No se pudo sincronizar el botón Continuar Alta:", e);
    }
  }



  // Normalizador compatible nuevo/legacy (incluye duplicado, referencia, PAN y tipo de cambio + aclaración + __raw)
  private normalizarMovimientos(apiBody: any, categoriaActual: string = "") {
    const out: any[] = [];
    try {
      const esNueva =
        Array.isArray(apiBody?.movimientos) &&
        apiBody.movimientos.length > 0 &&
        !apiBody.movimientos[0]?.transacciones;

      if (esNueva) {
        const arr = apiBody.movimientos as any[];

        // Filtramos por tipoTransaccion según categoría
        const tipoEsperado =
          categoriaActual === "Cuentas" ? 2 :
          (categoriaActual === "Tarjeta de crédito" || categoriaActual === "Tarjeta de débito") ? 1 : undefined;

        const filtrados = arr.filter(x => {
          if (tipoEsperado == null) return true;
          const t = Number(x?.tipoTransaccion);
          return t === tipoEsperado;
        });

        for (const t of filtrados) {
          const comercio = t?.nombreComercio ?? t?.descripcionOperacion ?? "";
          const fechaISO = t?.fechaOperacion ?? t?.fechaAutorizacion ?? null;
          const sign = (t?.indicadorCargoAbono === "-") ? -1 : 1;
          const monto = Number(t?.importe?.monto ?? t?.importeOriginal?.monto ?? t?.montoOriginal?.monto ?? 0) * sign;

          const referencia = t?.numeroReferencia ?? t?.referencia ?? "";
          const autorizacion = t?.autorizacion ?? t?.numeroAutorizacion ?? "";
          const panCrudo = t?.panTarjeta ?? t?.numeroTarjeta ?? "";
          const pan = this.maskCard(panCrudo);
          const divisaOriginal = t?.montoOriginal?.divisa ?? t?.descripcionMonedaOriginal ?? "";
          const tc = t?.tipoCambio ?? t?.tipoDeCambio ?? "";
          const tipoCambio = divisaOriginal && divisaOriginal !== "MXN"
            ? (tc ? `${divisaOriginal} · TC ${tc}` : `${divisaOriginal}`)
            : (tc ? `TC ${tc}` : "");

          const codigoFactura = t?.factura ?? t?.codigoFactura ?? t?.ticket ?? "";

          const duplicado = this.asBool(t?.movimientoDuplicado);
          const aclaracion =
            t?.estatusAclaracion ??
            t?.estatusAclaracionDescripcion ??
            t?.estatusAclaracionEstado ??
            t?.estatusAclaracionEstatus ??
            t?.aclaracionEstatus ??
            t?.aclaracion?.estatus ??
            undefined;

          out.push({
            comercio: String(comercio || ""),
            fechaISO: fechaISO || "",
            fecha: fechaISO ? this.toDateStd(fechaISO) : "",
            monto: isFinite(monto) ? monto : undefined,
            referencia: String(referencia || ""),
            autorizacion: String(autorizacion || ""),
            pan: pan || "",
            tipoCambio: String(tipoCambio || ""),
            codigoFactura: String(codigoFactura || ""),
            duplicado: duplicado,
            aclaracion,
            __raw: t,
          });
        }
      } else {
        // Legacy
        const lista = Array.isArray(apiBody?.movimientos) ? apiBody.movimientos : [];
        const match = lista[0];
        const trans = Array.isArray(match?.transacciones) ? match.transacciones : [];
        for (const t of trans) {
          const fechaISO = t?.fecha ? `${t?.fecha}T${t?.hora ?? "00:00:00"}` : "";
          out.push({
            comercio: t?.comercio ?? "",
            fechaISO,
            fecha: fechaISO ? this.toDateStd(fechaISO) : "",
            monto: isFinite(Number(t?.monto)) ? Number(t?.monto) : undefined,
            referencia: t?.referencia ?? "",
            autorizacion: t?.autorizacion ?? "",
            pan: this.maskCard(t?.numeroTarjeta),
            tipoCambio: t?.tipoCambio ? `TC ${t?.tipoCambio}` : "",
            codigoFactura: String(t?.factura ?? t?.codigoFactura ?? ""),
            duplicado: this.asBool(t?.movimientoDuplicado),
            aclaracion: t?.estatusAclaracion ?? undefined,
            __raw: t,
          });
        }
      }
    } catch (e) {
      console.error("[PCF][Movs] Error normalizando:", e);
    }

    // index estable para selección
    return out.map((r, i) => ({ ...r, __rowIndex: r.__rowIndex ?? i }));
  }

  /**
   * Dado un movimiento normalizado (row), intenta inferir el "tipo"
   * usando la matriz de códigos this.movTipoPorCodigo.
   */
  private inferTipoCodigoParaMovimiento(row: any): string | null {
    if (!row) return null;

    const raw = (row.codigoFactura ?? "").toString().trim();
    if (!raw) return null;

    const digits = raw.replace(/\D+/g, "");
    const tipo =
      this.movTipoPorCodigo[raw] ||
      (digits ? this.movTipoPorCodigo[digits] : undefined);

    return tipo ?? null;
  }

  /**
   * Determina si un movimiento es permitido según la matriz de códigos:
   * - Si NO hay matriz configurada (sin códigos) => no restringimos.
   * - Si SÍ hay matriz:
   *   - Solo se permite seleccionar si el código de factura del movimiento
   *     está dentro de la matriz de códigos de la configuración.
   *   - Siempre validamos contra la UNIÓN de todos los códigos (movCodigosMatrizAll / movConfigDebug.codigos).
   *   - Adicionalmente, si ya hay un tipo fijado (movTipoFijado), restringimos
   *     a los códigos de ese tipo.
   */
  private isMovimientoPermitidoPorMatriz(row: any): boolean {
    if (!row) return false;

    const codigoRaw = (row.codigoFactura ?? "").toString().trim();
    if (!codigoRaw) return false;

    const normalize = (v: string) => {
      const raw = v.trim();
      const digits = raw.replace(/\D+/g, "");
      const noZeros = digits.replace(/^0+/, "") || digits;
      return { raw, digits, noZeros };
    };

    const rowNorm = normalize(codigoRaw);

    // === Matriz global: unión de todos los códigos configurados ===
    const matrizGlobal: string[] =
      (this.movCodigosMatrizAll && this.movCodigosMatrizAll.length
        ? this.movCodigosMatrizAll
        : (this.movConfigDebug?.codigos ?? [])) || [];

    // Si NO hay configuración de códigos → no restringimos nada
    if (!matrizGlobal.length) {
      return true;
    }

    const matchesList = (list: string[]): boolean => {
      for (const c of list) {
        const n = normalize((c ?? "").toString());
        if (!n.raw && !n.digits) continue;

        if (
          (n.raw && n.raw === rowNorm.raw) ||
          (n.digits && n.digits === rowNorm.digits) ||
          (n.noZeros && n.noZeros === rowNorm.noZeros)
        ) {
          return true;
        }
      }
      return false;
    };

    // 1) Primero: el movimiento DEBE pertenecer a la unión de códigos configurados
    const enMatrizGlobal = matchesList(matrizGlobal);
    if (!enMatrizGlobal) {
      // Este es exactamente el caso que quieres bloquear:
      // código de factura que NO está en la matriz de configuración.
      return false;
    }

    // 2) Si aún NO hay tipo fijado (primer movimiento), cualquier código válido de la matriz global es seleccionable
    if (!this.movTipoFijado) {
      return true;
    }

    // 3) Si ya hay un tipo fijado, restringimos a los códigos de ese tipo
    const listaTipo: string[] =
      (this.movCodigosPorTipo &&
        this.movTipoFijado &&
        this.movCodigosPorTipo[this.movTipoFijado] &&
        this.movCodigosPorTipo[this.movTipoFijado].length
        ? this.movCodigosPorTipo[this.movTipoFijado]
        : (this.movCodigosMatriz || [])) || [];

    // Si por algún motivo no tenemos lista específica del tipo, dejamos solo la validación global
    if (!listaTipo.length) {
      return true;
    }

    const permitidoPorTipo = matchesList(listaTipo);

    return permitidoPorTipo;
  }





  /**
   * Carga la configuración de movimientos asociada a la Pregunta2 (o Pregunta1 si no hay Pregunta2)
   * y deja todo en this.movConfigDebug para poder inspeccionarlo en consola,
   * además de poblar:
   *  - this.movCodigosMatrizAll: unión de todos los códigos permitidos
   *  - this.movCodigosMatriz: lista activa (inicialmente SIN restricción; se llena al seleccionar el primer movimiento)
   *  - this.movCodigosPorTipo: mapa tipo -> códigos
   *  - this.movTipoPorCodigo: mapa código -> tipo
   */
  private async cargarConfigMovimientosParaSeleccion(): Promise<void> {
    try {
      // Reset SIEMPRE antes de cargar
      this.movConfigDebug = null;
      this.movCodigosMatriz = [];
      this.movCodigosMatrizAll = [];
      this.movCodigosPorTipo = {};
      this.movTipoPorCodigo = {};
      this.movTipoFijado = null;

      // Nuevo: también limpiamos el estado de P1/P2 movimiento
      this.state.movTipoSel = null;
      this.state.movP1Opciones = [];
      this.state.movP1SelId = null;
      this.state.movP2Opciones = [];
      this.state.movP2SelId = null;

      const p2 = this.state.preguntas2.find(p => p.id === this.state.preg2SelId) || null;
      const p1 = this.state.preguntas.find(p => p.id === this.state.preg1SelId) || null;

      const confMovIdRaw = p2?.confMovId || p1?.confMovId || null;
      const confMovId = this.cleanGuid(confMovIdRaw);
      if (!confMovId) {
        console.log("[PCF][MovCfg] Sin xmsbs_confmovimiento asociado a la selección actual.");
        return;
      }

      const api = this.getApi();
      if (!api?.retrieveMultipleRecords) {
        console.warn("[PCF][MovCfg] WebApi no disponible para cargar configuración de movimientos.");
        return;
      }

      // 1) Configuración principal
      const conf = await this.retrieveOne(
        "xmsbs_confmovimiento",
        "xmsbs_confmovimientoid,xmsbs_name",
        `xmsbs_confmovimientoid eq ${confMovId}`
      );

      // 2) Tipos de movimiento asociados a la conf
      const tiposRes = await api.retrieveMultipleRecords(
        "xmsbs_tipomovimiento",
        [
          "?$select=",
          [
            "xmsbs_tipomovimientoid",
            "xmsbs_name",
            "xmsbs_codigo",
            "xmsbs_ultimapregunta",
            "_xmsbs_subcategoria_value",
            "_xmsbs_confmovimiento_value",
          ].join(",")
          ,
          `&$filter=_xmsbs_confmovimiento_value eq ${confMovId}`
        ].join("")
      );

      const tipos: any[] = tiposRes?.entities ?? [];
      const tiposDebug: any[] = [];

      const matrizCodigos: string[] = [];

      for (const t of tipos) {
        const tipoId = this.cleanGuid(t.xmsbs_tipomovimientoid);
        const tipoName = t.xmsbs_name ?? "";
        const tipoCodigo = (t.xmsbs_codigo ?? "").toString().trim();
        const tipoSubId = t._xmsbs_subcategoria_value ?? null;
        const tipoUltima = this.asBool(t.xmsbs_ultimapregunta);

        tiposDebug.push({
          id: tipoId,
          name: tipoName,
          codigo: tipoCodigo,
          subcategoriaId: tipoSubId,
          ultimaPregunta: tipoUltima,
        });

        if (!tipoId) continue;

        // 3) Códigos asociados a cada tipo de movimiento
        const codRes = await api.retrieveMultipleRecords(
          "xmsbs_codigos",
          `?$select=xmsbs_codigosid,xmsbs_codigo,_xmsbs_tipomovimiento_value&$filter=_xmsbs_tipomovimiento_value eq ${tipoId}`
        );
        const cods: any[] = codRes?.entities ?? [];
        const codigosDeEsteTipo: string[] = [];

        for (const c of cods) {
          const raw = (c.xmsbs_codigo ?? "").toString().trim();
          if (!raw) continue;

          const norm = raw.replace(/\D+/g, "");
          const posibles: string[] = norm && norm !== raw ? [raw, norm] : [raw];

          for (const cod of posibles) {
            matrizCodigos.push(cod);
            codigosDeEsteTipo.push(cod);
            if (tipoCodigo) this.movTipoPorCodigo[cod] = tipoCodigo;
          }
        }

        if (tipoCodigo && codigosDeEsteTipo.length > 0) {
          this.movCodigosPorTipo[tipoCodigo] = Array.from(new Set(codigosDeEsteTipo));
        }
      }

      const matrizUnica = Array.from(new Set(matrizCodigos));
      this.movCodigosMatrizAll = matrizUnica;

      // NO aplicamos la matriz aún. Se activa cuando el usuario selecciona un movimiento.
      this.movCodigosMatriz = [];
      this.movTipoFijado = null;

      this.movConfigDebug = {
        confMovId,
        confMov: conf,
        tipos: tiposDebug,
        codigos: matrizUnica,
        codigosPorTipo: this.movCodigosPorTipo,
      };

      console.log("[PCF][MovCfg] Configuración de movimientos cargada:", this.movConfigDebug);
    } catch (e) {
      console.error("[PCF][MovCfg] Error al cargar configuración de movimientos:", e);
      this.movConfigDebug = null;
      this.movCodigosMatriz = [];
      this.movCodigosMatrizAll = [];
      this.movCodigosPorTipo = {};
      this.movTipoPorCodigo = {};
      this.movTipoFijado = null;

      // Limpieza de estado de P1/P2 movimiento
      this.state.movTipoSel = null;
      this.state.movP1Opciones = [];
      this.state.movP1SelId = null;
      this.state.movP2Opciones = [];
      this.state.movP2SelId = null;
    }
  }

  // Cargar P1Movimiento por Tipo de movimiento
  private async cargarP1MovimientosPorTipo(tipoMovId: string): Promise<void> {
    const api = this.getApi();
    const id = this.cleanGuid(tipoMovId);

    if (!api?.retrieveMultipleRecords || !id) {
      this.state.movP1Opciones = [];
      this.state.movP1SelId = null;
      this.state.movP2Opciones = [];
      this.state.movP2SelId = null;
      return;
    }

    const res = await api.retrieveMultipleRecords(
      "xmsbs_p1movimiento",
      [
        "?$select=",
        [
          "xmsbs_p1movimientoid",
          "xmsbs_name",
          "xmsbs_codigo",
          "xmsbs_ultimapregunta",
          "_xmsbs_subcategoria_value",
          "_xmsbs_tipomovimiento_value",
        ].join(","),
        `&$filter=_xmsbs_tipomovimiento_value eq ${id}`,
        "&$orderby=xmsbs_codigo asc",
      ].join("")
    );

    const lista: P1MovCfg[] = (res?.entities ?? []).map((e: any) => ({
      id: e.xmsbs_p1movimientoid,
      name: e.xmsbs_name ?? "",
      codigo: (e.xmsbs_codigo ?? "").toString(),
      subcategoriaId: e._xmsbs_subcategoria_value ?? null,
      ultimaPregunta: this.asBool(e.xmsbs_ultimapregunta),
    }));

    this.state.movP1Opciones = lista;
    this.state.movP1SelId = null;
    this.state.movP2Opciones = [];
    this.state.movP2SelId = null;
  }

  // Cargar P2Movimiento por P1Movimiento
  private async cargarP2MovimientosPorP1Mov(p1MovId: string): Promise<void> {
    const api = this.getApi();
    const id = this.cleanGuid(p1MovId);

    if (!api?.retrieveMultipleRecords || !id) {
      this.state.movP2Opciones = [];
      this.state.movP2SelId = null;
      return;
    }

    const res = await api.retrieveMultipleRecords(
      "xmsbs_p2movimiento",
      [
        "?$select=",
        [
          "xmsbs_p2movimientoid",
          "xmsbs_name",
          "xmsbs_codigo",
          "_xmsbs_p1mivimiento_value",
          "_xmsbs_subcategoria_value",
        ].join(","),
        `&$filter=_xmsbs_p1mivimiento_value eq ${id}`,
        "&$orderby=xmsbs_codigo asc",
      ].join("")
    );

    const lista: P2MovCfg[] = (res?.entities ?? []).map((e: any) => ({
      id: e.xmsbs_p2movimientoid,
      name: e.xmsbs_name ?? "",
      codigo: (e.xmsbs_codigo ?? "").toString(),
      p1MovId: e._xmsbs_p1mivimiento_value,
      subcategoriaId: e._xmsbs_subcategoria_value ?? null,
    }));

    this.state.movP2Opciones = lista;
    this.state.movP2SelId = null;
  }

  /**
   * Cuando cambia la selección de movimientos:
   * - Detecta el tipo de movimiento (xmsbs_tipomovimiento) en base al código.
   * - Si ese tipo tiene ultimaPregunta = sí → usamos su subcategoría directamente.
   * - Si ultimaPregunta = no → cargamos P1Movimiento para ese tipo.
   */
  private async onMovSelectionChanged(lastRow: any | null): Promise<void> {
    try {
      const selSize = this.state.movTable.selected?.size ?? 0;

      if (!this.movConfigDebug) {
        // Sin configuración, no aplicamos lógica nueva
        this.state.movTipoSel = null;
        this.state.movP1Opciones = [];
        this.state.movP1SelId = null;
        this.state.movP2Opciones = [];
        this.state.movP2SelId = null;
        this.recomputeFinalizarPorMovimientos();
        this.refreshMovTableUI();
        return;
      }

      if (selSize === 0) {
        // Nada seleccionado → limpiamos todo
        this.state.movTipoSel = null;
        this.state.movP1Opciones = [];
        this.state.movP1SelId = null;
        this.state.movP2Opciones = [];
        this.state.movP2SelId = null;
        this.recomputeFinalizarPorMovimientos();
        this.refreshMovTableUI();
        this.render();
        return;
      }

      // Tomamos un movimiento representativo (el clickeado si sigue seleccionado, si no el primero del set)
      let row = lastRow && this.state.movTable.selected.has(lastRow.__rowIndex) ? lastRow : null;
      if (!row) {
        const firstIdx = Array.from(this.state.movTable.selected)[0];
        row = (this.state.movimientos ?? []).find((m: any) => m.__rowIndex === firstIdx) || null;
      }
      if (!row) return;

      const tipoCodigo = this.inferTipoCodigoParaMovimiento(row);
      if (!tipoCodigo) {
        this.state.movTipoSel = null;
        this.state.movP1Opciones = [];
        this.state.movP1SelId = null;
        this.state.movP2Opciones = [];
        this.state.movP2SelId = null;
        this.recomputeFinalizarPorMovimientos();
        this.refreshMovTableUI();
        this.render();
        return;
      }

      const tiposCfg = (this.movConfigDebug?.tipos ?? []) as any[];
      const tipoInfo = tiposCfg.find(t => (t.codigo ?? "").toString().trim() === tipoCodigo) || null;
      if (!tipoInfo) {
        this.state.movTipoSel = null;
        this.state.movP1Opciones = [];
        this.state.movP1SelId = null;
        this.state.movP2Opciones = [];
        this.state.movP2SelId = null;
        this.recomputeFinalizarPorMovimientos();
        this.refreshMovTableUI();
        this.render();
        return;
      }

      this.state.movTipoSel = {
        id: this.cleanGuid(tipoInfo.id),
        name: tipoInfo.name ?? "",
        codigo: tipoInfo.codigo ?? "",
        subcategoriaId: tipoInfo.subcategoriaId ?? null,
        ultimaPregunta: this.asBool(tipoInfo.ultimaPregunta),
      };

      if (this.state.movTipoSel.ultimaPregunta) {
        // Caso 1 del diagrama: TipoMov = última → NO hay P1/P2Mov, tomamos subcategoría del tipo
        this.state.movP1Opciones = [];
        this.state.movP1SelId = null;
        this.state.movP2Opciones = [];
        this.state.movP2SelId = null;
      } else {
        // Caso 2: TipoMov no es última → cargamos P1Movimiento
        await this.cargarP1MovimientosPorTipo(this.state.movTipoSel.id);
      }

      this.recomputeFinalizarPorMovimientos();
      this.refreshMovTableUI();
      this.render();
    } catch (e) {
      console.error("[PCF][MovCfg] Error procesando selección de movimiento:", e);
    }
  }











  // ========= Helpers tabla: filtro/paginación =========
  private getFilteredMovs(): Array<any> {
    const rows = (this.state.movimientos ?? []).map((r, i) => ({ ...r, __rowIndex: r.__rowIndex ?? i }));
    const {
      searchText,
      filtroComercio,
      filtroReferencia,
      filtroAutorizacion,
      filtroPan,
      filtroTipoCambio,
      fechaDesde,
      fechaHasta,
      montoMin,
      montoMax,
      filtroDuplicados,
    } = this.state.movTable;

    const text = (searchText ?? "").trim().toLowerCase();

    // parse fechas
    const from = fechaDesde ? new Date(`${fechaDesde}T00:00:00`) : null;
    const to = fechaHasta ? new Date(`${fechaHasta}T23:59:59`) : null;

    const min = montoMin !== "" ? Number(montoMin) : null;
    const max = montoMax !== "" ? Number(montoMax) : null;

    const byText = (r: any) => {
      if (!text) return true;
      return [
        r.comercio,
        r.fecha,
        r.monto,
        r.referencia,
        r.autorizacion,
        r.pan,
        r.tipoCambio,
        r.duplicado ? "sí" : "no",
      ]
        .filter(Boolean)
        .some((v: string) => v.toString().toLowerCase().includes(text));
    };

    const byCols = (r: any) => {
      const okComercio = !filtroComercio || r.comercio?.toLowerCase().includes(filtroComercio.toLowerCase());
      const okRef = !filtroReferencia || r.referencia?.toLowerCase().includes(filtroReferencia.toLowerCase());
      const okAuto = !filtroAutorizacion || r.autorizacion?.toLowerCase().includes(filtroAutorizacion.toLowerCase());
      const okPan = !filtroPan || r.pan?.toLowerCase().includes(filtroPan.toLowerCase());
      const okTC = !filtroTipoCambio || r.tipoCambio?.toLowerCase().includes(filtroTipoCambio.toLowerCase());

      // Duplicados
      let okDup = true;
      if (filtroDuplicados === "con") okDup = !!r.duplicado;
      if (filtroDuplicados === "sin") okDup = !r.duplicado;

      // Rango fecha
      let okFecha = true;
      if (from || to) {
        const f = r.fechaISO ? new Date(r.fechaISO) : null;
        if (f) {
          if (from && f < from) okFecha = false;
          if (to && f > to) okFecha = false;
        }
      }

      // Rango monto
      let okMonto = true;
      if (min !== null && isFinite(min)) okMonto = (r.monto ?? -Infinity) >= min;
      if (okMonto && max !== null && isFinite(max)) okMonto = (r.monto ?? Infinity) <= max;

      return okComercio && okRef && okAuto && okPan && okTC && okDup && okFecha && okMonto;
    };

    return rows.filter(byText).filter(byCols);
  }

  private getPagedMovs(): { pageRows: any[]; total: number; start: number; end: number; lastPage: number } {
    const filtered = this.getFilteredMovs();
    const { pageIndex, pageSize } = this.state.movTable;
    const lastPage = Math.max(0, Math.ceil(filtered.length / pageSize) - 1);
    const safeIndex = Math.min(Math.max(0, pageIndex), lastPage);
    if (safeIndex !== pageIndex) this.state.movTable.pageIndex = safeIndex;

    const start = safeIndex * pageSize;
    const end = Math.min(filtered.length, start + pageSize);
    const pageRows = filtered.slice(start, end);
    return { pageRows, total: filtered.length, start, end, lastPage };
  }

  // ========= RENDER movimientos (tabla avanzada) =========
  // ========= RENDER movimientos (tabla avanzada) =========
  private renderMovimientos(): string {
    const { movLoading, movError, movimientos } = this.state;
    if (movLoading) {
      return `<div class="mov-container"><div class="muted">Cargando movimientos…</div></div>`;
    }
    if (movError) {
      return `<div class="mov-container"><div class="muted">${this.safe(movError)}</div></div>`;
    }
    if (!movimientos || movimientos.length === 0) {
      return `<div class="mov-container"><div class="muted">Sin movimientos para mostrar.</div></div>`;
    }

    const { pageRows, total, start, end } = this.getPagedMovs();
    const allSelectedOnPage = pageRows.every(r => this.state.movTable.selected.has(r.__rowIndex) || r.duplicado);
    const someSelectedOnPage = !allSelectedOnPage && pageRows.some(r => this.state.movTable.selected.has(r.__rowIndex));

    return `
        <div class="mov-container">
          <div class="section-header">Movimientos</div>

          <!-- Toolbar -->
          <div class="mov-toolbar">
            <div class="mov-toolbar-left">
              <div class="mov-search">
                <input id="mov-search" class="mov-input" type="text" autocomplete="off" placeholder="Buscar en todos los campos…" value="${this.safe(this.state.movTable.searchText)}" />
              </div>
            </div>
            <div class="mov-toolbar-right">
              <label class="mov-pagesize-label">Filas por página</label>
              <select id="mov-page-size" class="mov-select">
                ${[10,25,50,100].map(ps => `<option value="${ps}" ${this.state.movTable.pageSize===ps?"selected":""}>${ps}</option>`).join("")}
              </select>
              <div class="mov-pagination">
                <button class="mov-page-btn" data-mov-page="first" aria-label="Primera página">«</button>
                <button class="mov-page-btn" data-mov-page="prev" aria-label="Página anterior">‹</button>
                <span class="mov-page-status" id="mov-status-top">${start+1}-${end} de ${total}</span>
                <button class="mov-page-btn" data-mov-page="next" aria-label="Página siguiente">›</button>
                <button class="mov-page-btn" data-mov-page="last" aria-label="Última página">»</button>
              </div>
            </div>
          </div>

          <!-- Filtros (2 filas, 5 por fila) -->
          <div class="mov-filters">
            <div class="mov-filters-row">
              <input id="mov-filter-filtroComercio" class="mov-input" type="text" autocomplete="off" placeholder="Comercio / Descripción" value="${this.safe(this.state.movTable.filtroComercio)}" />
              <input id="mov-filter-filtroReferencia" class="mov-input" type="text" autocomplete="off" placeholder="Referencia" value="${this.safe(this.state.movTable.filtroReferencia)}" />
              <input id="mov-filter-filtroAutorizacion" class="mov-input" type="text" autocomplete="off" placeholder="Autorización" value="${this.safe(this.state.movTable.filtroAutorizacion)}" />
              <input id="mov-filter-filtroPan" class="mov-input" type="text" autocomplete="off" placeholder="PAN" value="${this.safe(this.state.movTable.filtroPan)}" />
              <input id="mov-filter-filtroTipoCambio" class="mov-input" type="text" autocomplete="off" placeholder="Tipo de cambio" value="${this.safe(this.state.movTable.filtroTipoCambio)}" />
            </div>
            <div class="mov-filters-row">
              <div class="mov-filter-range">
                <label>Fecha desde</label>
                <input id="mov-filter-fechaDesde" class="mov-input" type="date" value="${this.safe(this.state.movTable.fechaDesde)}" />
              </div>
              <div class="mov-filter-range">
                <label>Fecha hasta</label>
                <input id="mov-filter-fechaHasta" class="mov-input" type="date" value="${this.safe(this.state.movTable.fechaHasta)}" />
              </div>
              <div class="mov-filter-range">
                <label>Monto mín</label>
                <input id="mov-filter-montoMin" class="mov-input" type="number" step="0.01" inputmode="decimal" value="${this.safe(this.state.movTable.montoMin)}" />
              </div>
              <div class="mov-filter-range">
                <label>Monto máx</label>
                <input id="mov-filter-montoMax" class="mov-input" type="number" step="0.01" inputmode="decimal" value="${this.safe(this.state.movTable.montoMax)}" />
              </div>
              <div class="mov-filter-range">
                <label>Duplicados</label>
                <select id="mov-filter-filtroDuplicados" class="mov-select">
                  <option value="todos" ${this.state.movTable.filtroDuplicados==='todos'?'selected':''}>Todos</option>
                  <option value="con" ${this.state.movTable.filtroDuplicados==='con'?'selected':''}>Con aclaración</option>
                  <option value="sin" ${this.state.movTable.filtroDuplicados==='sin'?'selected':''}>Sin aclaración</option>
                </select>
              </div>
            </div>
          </div>

          <!-- Tabla -->
          <div class="mov-table-wrapper">
            <table class="mov-table" role="table" aria-label="Tabla de movimientos">
              <thead>
                <tr role="row">
                  <th class="sel-col" role="columnheader" aria-label="seleccionar">
                    <fluent-checkbox id="mov-select-all" disabled></fluent-checkbox>
                  </th>
                  <th role="columnheader">Comercio / Descripción</th>
                  <th role="columnheader">Importe</th>
                  <th role="columnheader">Fecha y hora de la operación</th>
                  <th role="columnheader">Número de referencia</th>
                  <th role="columnheader">Número de autorización</th>
                  <th role="columnheader">PAN de la tarjeta</th>
                  <th role="columnheader">Factura</th>
                  <th role="columnheader">Tipo de cambio</th>
                  <th role="columnheader">Duplicado</th>
                </tr>
              </thead>
              <tbody id="mov-tbody">
                ${this.buildMovTbodyRows(pageRows)}
              </tbody>
            </table>
          </div>

          <!-- Paginación inferior -->
          <div class="mov-toolbar mov-toolbar-bottom">
            <div class="mov-toolbar-right">
              <div class="mov-pagination">
                <button class="mov-page-btn" data-mov-page="first" aria-label="Primera página">«</button>
                <button class="mov-page-btn" data-mov-page="prev" aria-label="Página anterior">‹</button>
                <span class="mov-page-status" id="mov-status-bottom">${start+1}-${end} de ${total}</span>
                <button class="mov-page-btn" data-mov-page="next" aria-label="Página siguiente">›</button>
                <button class="mov-page-btn" data-mov-page="last" aria-label="Última página">»</button>
              </div>
            </div>
          </div>

          ${this.renderMovimientosExtras()}
        </div>
      `;
}



  // Bloque extra: combos de P1Movimiento / P2Movimiento debajo de la tabla
  private renderMovimientosExtras(): string {
    const tipo = this.state.movTipoSel;
    const p1List = this.state.movP1Opciones ?? [];
    const p2List = this.state.movP2Opciones ?? [];

    // Si no hay tipo o no hay nada que mostrar, salimos
    if (!tipo && (!p1List.length && !p2List.length)) return "";

    const selectedP1 = p1List.find(m => m.id === this.state.movP1SelId) || null;
    const esUltimaP1 = selectedP1 ? this.asBool(selectedP1.ultimaPregunta) : false;

    const showP1 = !!tipo && !tipo.ultimaPregunta && p1List.length > 0;
    const showP2 = showP1 && !esUltimaP1 && p2List.length > 0;

    if (!showP1 && !showP2) return "";

    const p1Html = showP1
      ? `
        <div class="mov-extra-row">
          <label class="question-label" id="lblP1Mov">¿P1 Movimiento?</label>
          <div class="question-value">
            <select id="mov-p1-select" class="mov-select">
              <option value="">-- Selecciona P1 movimiento --</option>
              ${p1List
                .map(
                  (m) => `
                    <option value="${this.safe(m.id)}" ${this.state.movP1SelId === m.id ? "selected" : ""}>
                      ${this.safe(m.name)}
                    </option>`
                )
                .join("")}
            </select>
          </div>
        </div>`
      : "";

    const p2Html = showP2
      ? `
        <div class="mov-extra-row">
          <label class="question-label" id="lblP2Mov">¿P2 Movimiento?</label>
          <div class="question-value">
            <select id="mov-p2-select" class="mov-select">
              <option value="">-- Selecciona P2 movimiento --</option>
              ${p2List
                .map(
                  (m) => `
                    <option value="${this.safe(m.id)}" ${this.state.movP2SelId === m.id ? "selected" : ""}>
                      ${this.safe(m.name)}
                    </option>`
                )
                .join("")}
            </select>
          </div>
        </div>`
      : "";

    return `
      <div class="mov-extra-container">
        ${p1Html}
        ${p2Html}
      </div>
    `;
  }


  private buildMovTbodyRows(rows: any[]): string {
    return rows
      .map((r) => {
        const selected = this.state.movTable.selected.has(r.__rowIndex);
        const permitidoMatriz = this.isMovimientoPermitidoPorMatriz(r);
        const disabled = !!r.duplicado || !permitidoMatriz; // no seleccionable
        const tooltip = this.safe(
          r.aclaracion ||
            (r.duplicado ? "Aclaración en proceso" : "") ||
            (!permitidoMatriz ? "Movimiento no permitido por la matriz de códigos" : "")
        );
        return `
          <tr role="row" class="${selected ? "row-selected" : ""}">
            <td class="sel-col" role="gridcell">
              <div class="mov-row-checkbox" data-row-index="${r.__rowIndex}">
                <fluent-checkbox ${selected ? "checked" : ""} ${disabled ? "disabled" : ""}></fluent-checkbox>
              </div>
            </td>
            <td role="gridcell">${this.safe(r.comercio)}</td>
            <td role="gridcell">${this.formatMoney(r.monto)}</td>
            <td role="gridcell">${this.safe(r.fecha)}</td>
            <td role="gridcell">${this.safe(r.referencia)}</td>
            <td role="gridcell">${this.safe(r.autorizacion)}</td>
            <td role="gridcell">${this.safe(r.pan)}</td>
            <td role="gridcell">${this.safe(r.codigoFactura)}</td>
            <td role="gridcell">${this.safe(r.tipoCambio)}</td>
            <td role="gridcell">
              ${
                r.duplicado
                  ? `<span class="dup-badge dup-yes" title="${this.safe(tooltip)}">Duplicado</span>`
                  : `<span class="dup-badge dup-no" title="${this.safe(tooltip)}">No</span>`
              }
            </td>
          </tr>
        `;
      })
      .join("");
  }



  // ======== REFRESCO PARCIAL DE TABLA ========
  private refreshMovTableUI() {
    const tbody = this.container.querySelector("#mov-tbody") as HTMLTableSectionElement | null;
    if (!tbody) return;

    const { pageRows, total, start, end } = this.getPagedMovs();

    // Reemplazamos solo el cuerpo
    tbody.innerHTML = this.buildMovTbodyRows(pageRows);

    // Actualizamos estado de “select all”
    const allSelectedOnPage = pageRows.every(r => this.state.movTable.selected.has(r.__rowIndex) || r.duplicado);
    const someSelectedOnPage = !allSelectedOnPage && pageRows.some(r => this.state.movTable.selected.has(r.__rowIndex));
    const selectAll = this.container.querySelector("#mov-select-all") as any;
    if (selectAll) {
      selectAll.checked = allSelectedOnPage;
      try { selectAll.indeterminate = someSelectedOnPage; } catch {}
      if (someSelectedOnPage) selectAll.setAttribute("indeterminate", "");
      else selectAll.removeAttribute("indeterminate");
    }

    // Contadores de paginación
    const top = this.container.querySelector("#mov-status-top");
    const bot = this.container.querySelector("#mov-status-bottom");
    if (top) top.textContent = `${start + 1}-${end} de ${total}`;
    if (bot) bot.textContent = `${start + 1}-${end} de ${total}`;
  }

  // ========= RELLENO LOOKUPS EN FORM =========
  private setFormLookup(logicalName: string, id: string, name: string, entityType: string) {
    try {
      const XrmAny = (window as any).Xrm;
      const ctx = XrmAny?.Page ?? XrmAny?.getFormContext?.();
      const attr = ctx?.getAttribute?.(logicalName);
      if (!attr?.setValue) return;
      attr.setValue([{ id: `{${id.replace(/[{}]/g,"")}}`, name: name || "", entityType }]);
      attr.fireOnChange?.();
    } catch (e) {
      console.warn("No se pudo setear lookup", logicalName, e);
    }
  }

  private async retrieveOne(entity: string, select: string, filter: string): Promise<any|null> {
    const api = this.getApi();
    if (!api?.retrieveMultipleRecords) return null;
    const query = `?$select=${select}${filter ? `&$filter=${filter}` : ""}`;
    const res = await api.retrieveMultipleRecords(entity, query);
    return (res?.entities && res.entities[0]) ? res.entities[0] : null;
  }

  private async getSubcategoriaDetalles(subcatId: string): Promise<{
    categoria?: { id: string, name: string, entityType: string },
    flujo?: { id: string, name: string, entityType: string },
  }> {
    const api = this.getApi();
    const id = this.cleanGuid(subcatId);
    if (!api?.retrieveMultipleRecords || !id) return {};

    const sub = await this.retrieveOne(
      "xmsbs_subcategoria",
      "xmsbs_subcategoriaid,xmsbs_name,_xmsbs_categoria_value,_xmsbs_flujo_value",
      `xmsbs_subcategoriaid eq ${id}`
    );

    const out: any = {};
    if (sub?._xmsbs_categoria_value) {
      out.categoria = {
        id: sub._xmsbs_categoria_value,
        name: sub["_xmsbs_categoria_value@OData.Community.Display.V1.FormattedValue"] || "",
        entityType: "xmsbs_categoria",
      };
    }
    if (sub?._xmsbs_flujo_value) {
      out.flujo = {
        id: sub._xmsbs_flujo_value,
        name: sub["_xmsbs_flujo_value@OData.Community.Display.V1.FormattedValue"] || "",
        entityType: "xmsbs_flujo",
      };
    }
    return out;
  }

  private async getEtapaInicial(flujoId?: string): Promise<{id:string,name:string,entityType:string}|null> {
    if (!flujoId) return null;
    const flu = this.cleanGuid(flujoId);
    if (!flu) return null;

    const e = await this.retrieveOne(
      "xmsbs_etapa",
      "xmsbs_etapaid,xmsbs_name,xmsbs_orden,_xmsbs_flujo_value",
      `_xmsbs_flujo_value eq ${flu} and xmsbs_orden eq 1`
    );
    if (!e?.xmsbs_etapaid) return null;
    return { id: e.xmsbs_etapaid, name: e.xmsbs_name ?? "", entityType: "xmsbs_etapa" };
  }

  private async rellenarLookupsCaso(): Promise<void> {
    try {
      // === 1) PRODUCTO (igual que antes) ===
      const codigo = this.categoriaToCodigo[this.state.categoria];
      const producto = codigo ? await this.getProductoPorCodigo(codigo) : null;
      if (producto) {
        this.setFormLookup("xmsbs_producto", producto.id, producto.name, producto.entityType);
      }

      // Preguntas 1 y 2 (para mantener los lookups viejos en el caso)
      const p1 = this.state.preguntas.find(p => p.id === this.state.preg1SelId) || null;
      const p2 = this.state.preguntas2.find(p => p.id === this.state.preg2SelId) || null;

      if (p1?.id) {
        this.setFormLookup("xmsbs_pregunta1", p1.id, p1.name || "", "xmsbs_pregunta1");
      }
      if (p2?.id) {
        this.setFormLookup("xmsbs_pregunta2", p2.id, p2.name || "", "xmsbs_pregunta2");
      }

      // === 2) NUEVO: movimiento seleccionado (TipoMov / P1Mov / P2Mov) ===
      const tipoMovSel = this.state.movTipoSel || null;
      const p1MovSel =
        (this.state.movP1Opciones ?? []).find(m => m.id === this.state.movP1SelId) || null;
      const p2MovSel =
        (this.state.movP2Opciones ?? []).find(m => m.id === this.state.movP2SelId) || null;

      // Siempre que tengamos tipo/p1/p2 de movimiento, llenamos esos lookups dedicados
      if (tipoMovSel?.id) {
        this.setFormLookup(
          "xmsbs_tipodemovimiento",
          tipoMovSel.id,
          tipoMovSel.name || "",
          "xmsbs_tipomovimiento"
        );
      }

      if (p1MovSel?.id) {
        this.setFormLookup(
          "xmsbs_pregunta1movimiento",
          p1MovSel.id,
          p1MovSel.name || "",
          "xmsbs_p1movimiento"
        );
      }

      if (p2MovSel?.id) {
        this.setFormLookup(
          "xmsbs_pregunta2movimiento",
          p2MovSel.id,
          p2MovSel.name || "",
          "xmsbs_p2movimiento"
        );
      }

      // === 3) NUEVO: PRIORIDAD para la SUBCATEGORÍA del caso ===
      //
      // Regla:
      //   1) Si hay escenario de movimientos, intentamos sacar la subcategoría desde:
      //        a) TipoMov (si es últimaPregunta)
      //        b) P2Movimiento (si existe)
      //        c) P1Movimiento (si existe)
      //   2) Si no logramos nada con movimientos → caemos al comportamiento clásico:
      //        - Pregunta2.subcategoria
      //        - o Pregunta1.subcategoria
      //
      let subcatDesdeMov: string | null = null;

      // 3.1) Caso TipoMov = última pregunta → tomamos su subcategoría directamente
      if (tipoMovSel?.subcategoriaId && this.asBool(tipoMovSel.ultimaPregunta)) {
        subcatDesdeMov = tipoMovSel.subcategoriaId;
      }

      // 3.2) Si no fue por TipoMov, probamos con P2Movimiento (si existe y tiene subcategoría)
      if (!subcatDesdeMov && p2MovSel?.subcategoriaId) {
        subcatDesdeMov = p2MovSel.subcategoriaId;
      }

      // 3.3) Si tampoco, probamos con P1Movimiento
      if (!subcatDesdeMov && p1MovSel?.subcategoriaId) {
        subcatDesdeMov = p1MovSel.subcategoriaId;
      }

      // 3.4) Definimos la subcategoría FINAL del caso:
      //      primero la que venga de movimientos (si hay),
      //      si no, la que ya venías usando desde P2/P1.
      let subcatFinal: string | null = null;

      if (subcatDesdeMov) {
        subcatFinal = subcatDesdeMov;
      } else if (p2?.subcategoriaId) {
        subcatFinal = p2.subcategoriaId || null;
      } else if (p1?.subcategoriaId) {
        subcatFinal = p1.subcategoriaId || null;
      }

      // 3.5) Con la subcategoría final, rellenamos:
      //      - xmsbs_subcategoria
      //      - xmsbs_categoria
      //      - xmsbs_flujo
      //      - xmsbs_etapa (etapa inicial del flujo)
      if (subcatFinal) {
        // Subcategoría
        this.setFormLookup("xmsbs_subcategoria", subcatFinal, "", "xmsbs_subcategoria");

        // Detalles de subcategoría → categoría y flujo
        const det = await this.getSubcategoriaDetalles(subcatFinal);
        if (det.categoria) {
          this.setFormLookup(
            "xmsbs_categoria",
            det.categoria.id,
            det.categoria.name,
            det.categoria.entityType
          );
        }
        if (det.flujo) {
          this.setFormLookup(
            "xmsbs_flujo",
            det.flujo.id,
            det.flujo.name,
            det.flujo.entityType
          );
        }

        // Etapa inicial del flujo (si hay flujo)
        const etapa = await this.getEtapaInicial(det.flujo?.id);
        if (etapa) {
          this.setFormLookup("xmsbs_etapa", etapa.id, etapa.name, etapa.entityType);
        }
      }

      // Resumen del nuevo comportamiento:
      // - Si el caso pasó por la lógica de movimientos, la subcategoría/categoría/flujo/etapa
      //   se decide en base a TipoMov / P1Mov / P2Mov (según corresponda).
      // - Si no hubo movimientos (o no hay config), sigue funcionando como antes
      //   usando la subcategoría de Pregunta2 o Pregunta1.

    } catch (e) {
      console.warn("[PCF] No se pudieron rellenar los lookups del caso:", e);
    }
  }




  // ========= CREAR xmsbs_contrato RELACIONADO (helper de guardado) =========

  private async saveCaseNow(): Promise<void> {
    try {
      const XrmAny = (window as any).Xrm;
      const ctx = XrmAny?.Page ?? XrmAny?.getFormContext?.();
      const saveFn =
        ctx?.data?.save?.bind(ctx?.data) ??
        XrmAny?.getFormContext?.()?.data?.save?.bind(XrmAny?.getFormContext?.()?.data);
      if (typeof saveFn === "function") {
        await saveFn();
      }
    } catch (e) {
      console.warn("[PCF] No se pudo guardar el Caso antes de crear contrato:", e);
    }
  }

  private getCurrentCaseId(): string | null {
    try {
      const XrmAny = (window as any).Xrm;
      const ctx = XrmAny?.Page ?? XrmAny?.getFormContext?.();
      const id = ctx?.data?.entity?.getId?.() || ctx?.data?.getEntity?.()?.getId?.();
      if (!id) return null;
      return id.replace(/[{}]/g, "").toLowerCase();
    } catch {
      return null;
    }
  }

  // --- FECHAS PARA DATAVERSE (ISO y válidas >= 1753-01-01) ---
  private readonly MIN_CRM_DATE_UTC = new Date(Date.UTC(1753, 0, 1, 0, 0, 0));

  private toDateIso(d: any): string | undefined {
    if (!d) return undefined;
    const dt = new Date(d);
    if (isNaN(dt.getTime())) return undefined;
    if (dt < this.MIN_CRM_DATE_UTC) return undefined;
    return dt.toISOString();
  }

  private fromYearMonth(y?: any, m?: any): string | undefined {
    const yy = Number(y), mm = Number(m);
    if (!isFinite(yy) || !isFinite(mm) || yy < 1753 || mm < 1 || mm > 12) return undefined;
    const dt = new Date(Date.UTC(yy, mm - 1, 1, 0, 0, 0));
    if (dt < this.MIN_CRM_DATE_UTC) return undefined;
    return dt.toISOString();
  }

  // IMPORTANTE: omitir fechas/campos nulos. Ya devolvemos undefined cuando algo no parsea.
  private buildContratoPayloadFromProducto(prod: any): Record<string, any> {
    const raw = prod?.raw ?? {};
    const tipo: string = prod?.tipo ?? "";

    const base: Record<string, any> = {
      xmsbs_name: `${tipo} · ${prod?.contratoId ?? prod?.productoId ?? ""}`.trim(),
      xmsbs_centroalta: this.toWhole(raw?.centroAlta),
      xmsbs_clabe: this.toWhole(raw?.clabe),
      xmsbs_codigoproductosubproducto: (raw?.producto?.codigo && raw?.subproducto?.codigo)
        ? `${raw.producto.codigo}-${raw.subproducto.codigo}` : undefined,
      xmsbs_producto: raw?.producto?.descripcion ? undefined : raw?.producto?.codigo,
      xmsbs_productodescripcion: raw?.producto?.descripcion,
      xmsbs_subproducto: raw?.subproducto?.codigo,
      xmsbs_saldodisponible: this.toCurrency(raw?.saldoDisponible?.monto),
      xmsbs_fechaaltacontrato: this.toDateIso(raw?.fechaAltaContrato),
      xmsbs_estatus: this.toWhole(
        raw?.estatusTarjeta?.codigoInt ??
        raw?.estatusCredito?.codigoInt ??
        raw?.estadoInversion?.codigoInt ??
        raw?.estadoSeguro?.codigoInt
      ),
      xmsbs_tipotasa: raw?.tipoTasa || undefined,
      xmsbs_restricciones: raw?.restricciones || undefined,
    };

    if (tipo === "Tarjeta de Crédito") {
      Object.assign(base, {
        xmsbs_limitecredito: this.toCurrency(raw?.limiteCredito?.monto),
        xmsbs_montoapagartotal: this.toCurrency(raw?.montoPagarTotal?.monto),
        xmsbs_montoapagarminimo: this.toCurrency(raw?.montoPagarMinimo?.monto),
        xmsbs_fechacorte: this.toDateIso(raw?.fechaCorte),
        xmsbs_fechapago: this.toDateIso(raw?.fechaPago),
        xmsbs_fechabloqueo: this.toDateIso(raw?.fechaBloqueo),
        xmsbs_fechaactivaciontarjeta: this.toDateIso(raw?.fechaActivacionTarjeta),
        xmsbs_fechavencimiento: this.fromYearMonth(raw?.fechaVencimientoYear, raw?.fechaVencimientoMonth),
        xmsbs_tipobloqueo: this.toWhole(raw?.tipoBloqueo?.codigoInt),
        xmsbs_tipotarjeta: this.toWhole(raw?.tipoTarjeta?.codigoInt),
        xmsbs_indicadormarca: this.toWhole(raw?.indicadorMarca?.codigoInt),
        xmsbs_indicadorparticipacion: raw?.indicadorParticipacion?.descripcion || undefined,
        xmsbs_numerotarjeta: this.safeLastDigitsAsWhole(raw?.numeroTarjeta, 8),
        xmsbs_numerocredito: this.safeLastDigitsAsWhole(raw?.numeroContrato, 8),
      });
    } else if (tipo === "Tarjeta de Débito") {
      Object.assign(base, {
        xmsbs_fechabloqueo: this.toDateIso(raw?.fechaBloqueo),
        xmsbs_fechaactivaciontarjeta: this.toDateIso(raw?.fechaActivacionTarjeta),
        xmsbs_fechavencimiento: this.fromYearMonth(raw?.fechaVencimientoYear, raw?.fechaVencimientoMonth),
        xmsbs_tipobloqueo: this.toWhole(raw?.tipoBloqueo?.codigoInt),
        xmsbs_tipotarjeta: this.toWhole(raw?.tipoTarjeta?.codigoInt),
        xmsbs_indicadormarca: this.toWhole(raw?.indicadorMarca?.codigoInt),
        xmsbs_indicadorparticipacion: raw?.indicadorParticipacion?.descripcion || undefined,
        xmsbs_numerotarjeta: this.safeLastDigitsAsWhole(raw?.numeroTarjeta, 8),
        xmsbs_numerocredito: this.safeLastDigitsAsWhole(raw?.numeroContrato, 8),
      });
    } else if (tipo === "Cuenta") {
      Object.assign(base, {});
    } else if (tipo === "Crédito") {
      Object.assign(base, {
        xmsbs_fechacorte: this.toDateIso(raw?.fechaCorte),
        xmsbs_fechapago: this.toDateIso(raw?.fechaPago),
        xmsbs_fechavencimiento: this.toDateIso(raw?.fechaVencimiento),
        xmsbs_montoapagartotal: this.toCurrency(raw?.montoPagarTotal?.monto),
        xmsbs_saldoinsoluto: this.toCurrency(raw?.saldoInsoluto?.monto),
        xmsbs_montooriginal: this.toCurrency(raw?.montoOriginal?.monto),
        xmsbs_plazo: this.toWhole(raw?.plazo?.monto),
        xmsbs_pagomensual: this.toCurrency(raw?.montoPagoMensual?.monto),
        xmsbs_tasainteres: this.percentToWhole(raw?.tasaInteres),
        xmsbs_estatuscreditos: this.toWhole(raw?.estatusCredito?.codigoInt),
        xmsbs_numerocredito: this.safeLastDigitsAsWhole(raw?.numeroContrato, 8),
      });
    } else if (tipo === "Inversión") {
      Object.assign(base, {
        xmsbs_montoinvertido: this.toCurrency(raw?.montoInvertido?.monto),
        xmsbs_rendimientoestimado: this.percentToWhole(raw?.rendimientoEstimadoTasaContratada),
        xmsbs_estadoactualinversion: this.toWhole(raw?.estadoInversion?.codigoInt),
      });
    } else if (tipo === "Seguro") {
      Object.assign(base, {
        xmsbs_poliza: raw?.numeroPoliza || undefined,
        xmsbs_canalcontratacion: this.toWhole(raw?.canalContratacion?.codigoInt),
        xmsbs_ramo: this.toWhole(raw?.ramo?.codigoInt),
        xmsbs_estadoactualseguros: this.toWhole(raw?.estadoSeguro?.codigoInt),
        xmsbs_fechavencimiento: this.toDateIso(raw?.fechaVencimiento),
      });
    }

    const payload: Record<string, any> = {};
    Object.entries(base).forEach(([k, v]) => {
      if (v !== undefined && v !== null && v !== "") payload[k] = v;
    });
    return payload;
  }

  // ========= NUEVO: construir payload de xmsbs_movimiento desde __raw y categoría =========
  private construirPayloadMovimientoDesdeRaw(rawMov: any, categoria: string): Record<string, any> {
    const p: Record<string, any> = {};
    const nz = (v:any) => (v==null || String(v).trim()==="" ? undefined : String(v).trim());

    // Campos comunes si existen
    const importe = rawMov?.importe?.monto ?? rawMov?.importeOriginal?.monto ?? rawMov?.montoOriginal?.monto;
    const fechaOp = rawMov?.fechaOperacion ?? rawMov?.fechaOperacionISO ?? rawMov?.fecha;
    const fechaAut = rawMov?.fechaAutorizacion ?? rawMov?.fechaHoraAutorizacion;

    const nroRef = rawMov?.numeroReferencia ?? rawMov?.referencia ?? rawMov?.transactionId;
    const nroAut = rawMov?.autorizacion ?? rawMov?.numeroAutorizacion;

    p["xmsbs_name"] = nz(nroRef) ?? nz(`${nroAut || ""}`.trim() || undefined);
    if (importe != null && isFinite(Number(importe))) p["xmsbs_importemonto"] = this.toCurrency(importe);

    const isoOp = this.toDateIso(fechaOp);
    if (isoOp) p["xmsbs_fechayhoraoperacion"] = isoOp;

    const isoAut = this.toDateIso(fechaAut);
    if (isoAut) p["xmsbs_fechayhoraautorizacion"] = isoAut;

    if (nroRef) p["xmsbs_name"] = String(nroRef);
    if (nroAut) p["xmsbs_autorizacion"] = String(nroAut);

    if (rawMov?.indicadorCargoAbono) p["xmsbs_indicadorcargoabono"] = String(rawMov.indicadorCargoAbono);

    const panIngreso = rawMov?.numeroTarjeta ?? rawMov?.panTarjeta;
    if (panIngreso) p["xmsbs_pandelatarjetaingreso"] = String(panIngreso);

    const panOperacion = rawMov?.panOperacion ?? rawMov?.panTransaccion ?? rawMov?.panTarjeta;
    if (panOperacion) p["xmsbs_pandelatarjetaoperacion"] = String(panOperacion);

    const descOp = rawMov?.descripcionOperacion ?? rawMov?.nombreComercio ?? rawMov?.comercio;
    if (descOp) p["xmsbs_descripciondelaoperacion"] = String(descOp);
    const nombreComercio = rawMov?.nombreComercio;
    if (nombreComercio) p["xmsbs_nombrecomercio"] = String(nombreComercio);

    const descMonedaOrig = rawMov?.descripcionMonedaOriginal ?? rawMov?.montoOriginal?.divisa;
    if (descMonedaOrig) p["xmsbs_descripcionmonedaoriginal"] = String(descMonedaOrig);
    const impDivOrig = rawMov?.montoOriginal?.monto;
    if (impDivOrig != null && isFinite(Number(impDivOrig))) p["xmsbs_importedivisaoriginal"] = this.toCurrency(impDivOrig);
    const tipoCambio = rawMov?.tipoCambio ?? rawMov?.tipoDeCambio;
    if (tipoCambio != null && isFinite(Number(tipoCambio))) p["xmsbs_tipodecambio"] = Number(tipoCambio);

    if (rawMov?.ubicacionComercio) p["xmsbs_ubicaciondelcomercio"] = String(rawMov.ubicacionComercio);

    const isTDC = categoria === "Tarjeta de crédito";
    const isTDD = categoria === "Tarjeta de débito";
    const isCTA = categoria === "Cuentas";

    if (isTDC) {
      if (rawMov?.afiliacion) p["xmsbs_afiliacion"] = String(rawMov.afiliacion);
      if (rawMov?.canalOperacion) p["xmsbs_canaldeoperacion"] = String(rawMov.canalOperacion);
      if (rawMov?.claveTransaccion) p["xmsbs_clavedetransaccion"] = String(rawMov.claveTransaccion);
      if (rawMov?.factura) p["xmsbs_factura"] = String(rawMov.factura);
      if (rawMov?.franquicia) p["xmsbs_franquicia"] = String(rawMov.franquicia);
      if (rawMov?.giroComercio) p["xmsbs_girocomercio"] = String(rawMov.giroComercio);
      if (rawMov?.modoEntrada) p["xmsbs_codigomododeentrada"] = String(rawMov.modoEntrada);
      if (rawMov?.nombreOrdenante) p["xmsbs_nombreordenante"] = String(rawMov.nombreOrdenante);
      if (rawMov?.saldoConAutorizacion?.monto != null) p["xmsbs_saldoconelqueseautorizolaoperacion"] = this.toCurrency(rawMov.saldoConAutorizacion.monto);
      if (rawMov?.estatus) p["xmsbs_estatus"] = String(rawMov.estatus);
      if (rawMov?.tokenC0) p["xmsbs_tokenc0"] = String(rawMov.tokenC0);
      if (rawMov?.tokenB2) { const d=this.toDateIso(rawMov.tokenB2); if (d) p["xmsbs_tokenb2"] = d; }
      if (rawMov?.tokenB3) p["xmsbs_tokenb3"] = String(rawMov.tokenB3);
      if (rawMov?.tokenB5) p["xmsbs_tokenb5"] = String(rawMov.tokenB5);
      if (rawMov?.tokenC4) p["xmsbs_tokenc4"] = String(rawMov.tokenC4);
      if (rawMov?.tokenPO) p["xmsbs_tokenpo"] = String(rawMov.tokenPO);
      if (rawMov?.tokenPY) p["xmsbs_tokenpy"] = String(rawMov.tokenPY);
      if (rawMov?.tokenQ2) p["xmsbs_tokenq2"] = String(rawMov.tokenQ2);
      if (rawMov?.tokenCO) p["xmsbs_tokenco"] = String(rawMov.tokenCO);
      if (rawMov?.productoM) p["xmsbs_productom"] = String(rawMov.productoM);
      if (rawMov?.indicadorComercioSeguro != null) p["xmsbs_indicadordecomercioseguro"] = this.toWhole(rawMov.indicadorComercioSeguro);
    }

    if (isTDD) {
      if (rawMov?.afiliacion) p["xmsbs_afiliacion"] = String(rawMov.afiliacion);
      if (rawMov?.canalOperacion) p["xmsbs_canaldeoperacion"] = String(rawMov.canalOperacion);
      if (rawMov?.claveTransaccion) p["xmsbs_clavedetransaccion"] = String(rawMov.claveTransaccion);
      if (rawMov?.descripcionOperacion) p["xmsbs_descripciondelaoperacion"] = String(rawMov.descripcionOperacion);
      if (rawMov?.factura) p["xmsbs_factura"] = String(rawMov.factura);
      if (rawMov?.franquicia) p["xmsbs_franquicia"] = String(rawMov.franquicia);
      if (rawMov?.giroComercio) p["xmsbs_girocomercio"] = String(rawMov.giroComercio);
      if (rawMov?.modoEntrada) p["xmsbs_codigomododeentrada"] = String(rawMov.modoEntrada);
      if (rawMov?.nombreComercio) p["xmsbs_nombrecomercio"] = String(rawMov.nombreComercio);
      if (rawMov?.saldoConAutorizacion?.monto != null) p["xmsbs_saldoconelqueseautorizolaoperacion"] = this.toCurrency(rawMov.saldoConAutorizacion.monto);
      if (rawMov?.estatus) p["xmsbs_estatus"] = String(rawMov.estatus);
      if (rawMov?.montoOriginal?.divisa) p["xmsbs_descripcionmonedaoriginal"] = String(rawMov.montoOriginal.divisa);
      if (rawMov?.montoOriginal?.monto != null) p["xmsbs_importedivisaoriginal"] = this.toCurrency(rawMov.montoOriginal.monto);
      if (rawMov?.tipoCambio != null && isFinite(Number(rawMov.tipoCambio))) p["xmsbs_tipodecambio"] = Number(rawMov.tipoCambio);
      if (rawMov?.tokenC0) p["xmsbs_tokenc0"] = String(rawMov.tokenC0);
      if (rawMov?.tokenB2) { const d=this.toDateIso(rawMov.tokenB2); if (d) p["xmsbs_tokenb2"] = d; }
      if (rawMov?.tokenB3) p["xmsbs_tokenb3"] = String(rawMov.tokenB3);
      if (rawMov?.tokenB5) p["xmsbs_tokenb5"] = String(rawMov.tokenB5);
      if (rawMov?.tokenC4) p["xmsbs_tokenc4"] = String(rawMov.tokenC4);
      if (rawMov?.tokenPO) p["xmsbs_tokenpo"] = String(rawMov.tokenPO);
      if (rawMov?.tokenPY) p["xmsbs_tokenpy"] = String(rawMov.tokenPY);
      if (rawMov?.tokenQ2) p["xmsbs_tokenq2"] = String(rawMov.tokenQ2);
      if (rawMov?.tokenCO) p["xmsbs_tokenco"] = String(rawMov.tokenCO);
    }

    if (isCTA) {
      if (rawMov?.claveTransaccion) p["xmsbs_clavedetransaccion"] = String(rawMov.claveTransaccion);
      if (rawMov?.indicadorCargoAbono) p["xmsbs_indicadorcargoabono"] = String(rawMov.indicadorCargoAbono);
      if (rawMov?.nombreBeneficiario) p["xmsbs_nombrebeneficiario"] = String(rawMov.nombreBeneficiario);
      if (rawMov?.nombreOrdenante) p["xmsbs_nombreordenante"] = String(rawMov.nombreOrdenante);
      if (rawMov?.numeroContratoIngreso) p["xmsbs_numerocontratotarjetaqueingresa"] = String(rawMov.numeroContratoIngreso);
    }

    const payload: Record<string, any> = {};
    Object.entries(p).forEach(([k,v]) => { if (v!==undefined && v!==null && v!=="") payload[k]=v; });
    return payload;
  }

  // ========= NUEVO: crear contrato y devolver Id (sin refresh hasta terminar) =========
  private async crearContratoRelacionadoYDevolverId(): Promise<string | null> {
    const prodSel = this.state.productoSel;
    if (!prodSel?.raw) {
      this.showCrmAlert("Primero selecciona un producto.");
      return null;
    }

    await this.rellenarLookupsCaso();
    await this.saveCaseNow();

    const caseId = this.getCurrentCaseId();
    if (!caseId) {
      this.showCrmAlert("No pude obtener el Id del Caso.");
      return null;
    }

    const payload = this.buildContratoPayloadFromProducto(prodSel);
    payload["xmsbs_caso@odata.bind"] = `/incidents(${caseId})`;

    const api = this.getApi();
    if (!api?.createRecord) throw new Error("WebApi no disponible");
    const res = await api.createRecord("xmsbs_contrato", payload);
    const contratoId: string | undefined = res?.id;

    return contratoId ?? null;
  }

  // ========= NUEVO: crear movimientos seleccionados =========
  private async crearMovimientosSeleccionados(contratoId: string, caseId: string): Promise<{ok:number, fail:number}> {
    const api = this.getApi();
    if (!api?.createRecord) throw new Error("WebApi no disponible");

    const sel = Array.from(this.state.movTable.selected ?? new Set<number>());
    const byIndex: Record<number, any> = {};
    for (const m of (this.state.movimientos ?? [])) byIndex[m.__rowIndex] = m;

    let ok = 0, fail = 0;
    for (const idx of sel) {
      const row = byIndex[idx];
      if (!row || row.duplicado) continue;

      const rawMov = row.__raw ?? {};
      const payload = this.construirPayloadMovimientoDesdeRaw(rawMov, this.state.categoria);

      payload["xmsbs_caso@odata.bind"] = `/incidents(${caseId})`;
      payload["xmsbs_contrato@odata.bind"] = `/xmsbs_contratos(${contratoId})`;

      try {
        // === AJUSTE 1: entidad correcta ===
        await api.createRecord("xmsbs_movimiento", payload);
        ok++;
      } catch (e) {
        console.error("[PCF] Error creando xmsbs_movimiento:", e, payload);
        fail++;
      }
    }
    return { ok, fail };
  }

  // ========= NUEVO: orquestador Continuar Alta =========
  private async continuarAlta(): Promise<void> {
    try {
      // (Reglas de habilitación siguen manejándose en la UI; aquí quitamos restricciones por movimientos)
      // === AJUSTE 2: SIN validar que haya movimientos cargados ni seleccionados ===

      // Crear contrato siempre
      this.showGlobalProgress("Creando contrato…");
      const contratoId = await this.crearContratoRelacionadoYDevolverId();
      if (!contratoId) {
        this.closeGlobalProgress();
        return;
      }

      const caseId = this.getCurrentCaseId();
      if (!caseId) {
        this.closeGlobalProgress();
        this.showCrmAlert("No pude obtener el Id del Caso para continuar.");
        return;
      }

      // Si el usuario seleccionó movimientos, crearlos; si no, continuar sin crearlos.
      let ok = 0, fail = 0;
      const selCount = (this.state.movTable?.selected?.size ?? 0);
      if (selCount > 0) {
        this.showGlobalProgress("Creando movimientos seleccionados…");
        const res = await this.crearMovimientosSeleccionados(contratoId, caseId);
        ok = res.ok; fail = res.fail;
      }

      // Refresh de formulario al final
      await this.refreshCurrentForm(caseId);
      this.closeGlobalProgress();

      if (selCount > 0) {
        this.showCrmAlert(`Contrato creado. Movimientos creados: ${ok}. Errores: ${fail}.`);
      } else {
        this.showCrmAlert(`Contrato creado. No se crearon movimientos (ninguno seleccionado).`);
      }
    } catch (e:any) {
      this.closeGlobalProgress();
      console.error("[PCF] Error en Continuar Alta:", e);
      this.showCrmAlert(`No se pudo completar la operación: ${e?.message ?? e}`);
    }
  }

  // ======== Utils ========
  private hasSubcategoriaEnCaso(): boolean {
    const v = this.getValueFromFormAttribute("xmsbs_subcategoria");
    try {
      if (!v) return false;
      const arr = typeof v === "string" ? JSON.parse(v) : v;
      const first = Array.isArray(arr) ? arr[0] : null;
      const id = (first?.id ?? "").toString().replace(/[{}]/g, "");
      return !!id && !/^0{8}-0{4}-0{4}-0{4}-0{12}$/i.test(id);
    } catch {
      return !!v;
    }
  }

  private setLoading(val: boolean) {
    const wasLoading = this.state.loading;
    this.state.loading = val;
    this.container.classList.toggle("loading", val);

    // Solo planificamos el autosave si:
    // - veníamos de un estado "loading"
    // - ya terminó la carga (!val)
    // - la API de persona ya se inició (apiStarted)
    // - aún no hicimos autosave
    // - y NO estamos en el caso "jsonPersona sin cambios"
    if (wasLoading && !val && this.apiStarted && !this.autoSaveDone && this.shouldAutoSaveAfterPersona) {
      this.planAutoSave().catch(console.error);
    }
  }


  private async planAutoSave() {
    await this.nextFrame();
    await this.nextFrame();
    await this.delay(400);
    await this.waitForHydration(5000, 200);
    await this.guardarCasoActual();
  }

  private nextFrame() {
    return new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
  }

  private async waitForHydration(timeoutMs = 5000, intervalMs = 200): Promise<void> {
    const start = Date.now();
    const checks: Array<[string, () => string]> = [
      ["xmsbs_jsonpersona", () => this.outJsonPersona],
      ["xmsbs_firstname", () => this.outFirstName],
      ["xmsbs_middlename", () => this.outMiddleName],
      ["xmsbs_lastname", () => this.outLastName],
    ];
    const equals = (a?: string | null, b?: string | null) =>
      (a ?? "").toString().trim() === (b ?? "").toString().trim();

    while (Date.now() - start < timeoutMs) {
      let allOk = true;
      for (const [logical, getExpected] of checks) {
        const expected = getExpected();
        if (!expected) continue;
        const current = this.getValueFromFormAttribute(logical);
        if (!equals(current, expected)) {
          allOk = false;
          break;
        }
      }
      if (allOk) return;
      await this.delay(intervalMs);
    }
  }

  private async guardarCasoActual(): Promise<void> {
    if (!this.apiStarted || this.autoSaveDone) return;

    this.notifyOutputChanged?.();
    await this.delay(150);

    try {
      this.showGlobalProgress();
      const XrmAny = (window as any).Xrm;
      const saveFn =
        XrmAny?.Page?.data?.save?.bind(XrmAny?.Page?.data) ??
        XrmAny?.getFormContext?.()?.data?.save?.bind(XrmAny?.getFormContext?.()?.data);

      if (typeof saveFn !== "function") {
        console.warn("[PCF] No se encontró API de guardado del formulario.");
        return;
      }

      await saveFn();
      this.autoSaveDone = true;
      console.log("[PCF] Caso guardado automáticamente tras obtener datos de la API.");
    } catch (err) {
      console.error("[PCF] Error al guardar el caso:", err);
    } finally {
      this.closeGlobalProgress();
    }
  }

  private async refreshCurrentForm(caseId: string) {
    try {
      const XrmAny = (window as any).Xrm;
      const ctx = XrmAny?.Page ?? XrmAny?.getFormContext?.();

      await (ctx?.data?.save?.() ?? Promise.resolve());
      if (ctx?.data?.refresh) await ctx.data.refresh(false);
      if (XrmAny?.Navigation?.openForm) {
        await XrmAny.Navigation.openForm({ entityName: "incident", entityId: caseId });
        return;
      }
      (window as any).location?.reload?.();
    } catch {
      try { (window as any).location?.reload?.(); } catch {}
    }
  }

  private delay(ms: number): Promise<void> {
    return new Promise<void>((resolve) => setTimeout(resolve, ms));
  }

  private addQueryParams(url: string, params: Record<string, string>) {
    const u = new URL(url);
    for (const [k, v] of Object.entries(params)) {
      if (v !== undefined && v !== null && v !== "") u.searchParams.set(k, v);
    }
    return u.toString();
  }

  private getValueFromFormAttribute(logicalName: string): string | null {
    try {
      const XrmAny = (window as any).Xrm;
      const attr =
        XrmAny?.Page?.getAttribute?.(logicalName) ||
        XrmAny?.getFormContext?.()?.getAttribute?.(logicalName);
      const v = attr?.getValue?.();
      if (v === null || v === undefined) return null;
      return typeof v === "string" ? v : JSON.stringify(v);
    } catch {
      return null;
    }
  }

  private setBound(prop: string, value: any) {
    const self: any = this as any;
    if (self[prop] !== value) {
      self[prop] = value;
      this.notifyOutputChanged?.();
    }
  }

  private nz(v: any): string {
    if (v === null || v === undefined) return "";
    if (typeof v === "number" && v === 0) return "";
    const s = String(v);
    return s === "0" ? "" : s;
  }

  private formatAntiguedad(a?: { annios?: number; meses?: number; dias?: number } | null): string {
    if (!a) return "";
    const parts: string[] = [];
    const y = (a.annios ?? 0) || 0;
    const m = (a.meses ?? 0) || 0;
    const d = (a.dias ?? 0) || 0;
    if (y > 0) parts.push(`${y} ${y === 1 ? "año" : "años"}`);
    if (m > 0) parts.push(`${m} ${m === 1 ? "mes" : "meses"}`);
    if (d > 0) parts.push(`${d} ${d === 1 ? "día" : "días"}`);
    return parts.join(", ");
  }

  private isEmptyGuid(id?: string): boolean {
    return !id || /^0{8}-0{4}-0{4}-0{4}-0{12}$/i.test(id);
  }

  private makeLookup(lkp: any): any {
    const id = lkp?.id as string | undefined;
    if (!id || this.isEmptyGuid(id)) return undefined;
    const name = lkp?.name ?? "";
    const entityType = lkp?.logicalName ?? lkp?.entityType ?? "";
    if (!entityType) return undefined;
    return [{ id, name, entityType }];
  }

  private toWhole(v: any): number | undefined {
    const n = Number(v);
    if (!isFinite(n)) return undefined;
    const i = Math.trunc(n);
    if (i > 2147483647 || i < -2147483648) return undefined;
    return i;
  }
  private toCurrency(v: any): number | undefined {
    const n = Number(v);
    return isFinite(n) ? n : undefined;
  }

  private safeLastDigitsAsWhole(value: any, take: number = 8): number | undefined {
    if (value === null || value === undefined) return undefined;
    const digits = String(value).replace(/\D+/g, "");
    if (!digits) return undefined;
    const last = digits.slice(-Math.max(1, take));
    const n = Number(last);
    if (!isFinite(n)) return undefined;
    if (n > 2147483647) return 2147483647;
    return Math.trunc(n);
  }

  private percentToWhole(v: any): number | undefined {
    const n0 = Number(v);
    if (!isFinite(n0)) return undefined;
    const n = Math.abs(n0) <= 1 ? (n0 * 100) : n0;
    const r = Math.round(n);
    if (r > 2147483647) return 2147483647;
    if (r < -2147483648) return -2147483648;
    return r;
  }

  // ==== ESTANDARIZAR FECHA A dd-mm-yyyy hh:mm:ss (SOLO UI) ====
  private toDateStd(d: any): string | undefined {
    if (!d) return undefined;
    const dt = new Date(d);
    if (isNaN(dt.getTime())) return undefined;
    const pad = (n: number) => String(n).padStart(2, "0");
    const dd = pad(dt.getDate());
    const MM = pad(dt.getMonth() + 1);
    const yyyy = dt.getFullYear();
    const hh = pad(dt.getHours());
    const mm = pad(dt.getMinutes());
    const ss = pad(dt.getSeconds());
    return `${dd}-${MM}-${yyyy} ${hh}:${mm}:${ss}`;
  }

  private fromYearMonthStd(y?: any, m?: any): string | undefined {
    const yy = Number(y), mm = Number(m);
    if (!isFinite(yy) || !isFinite(mm) || yy < 1900 || mm < 1 || mm > 12) return undefined;
    const dt = new Date(yy, mm - 1, 1, 0, 0, 0);
    const pad = (n: number) => String(n).padStart(2, "0");
    const dd = pad(dt.getDate());
    const MM = pad(dt.getMonth() + 1);
    const yyyy = dt.getFullYear();
    const hh = pad(dt.getHours());
    const mi = pad(dt.getMinutes());
    const ss = pad(dt.getSeconds());
    return `${dd}-${MM}-${yyyy} ${hh}:${mi}:${ss}`;
  }

  private getProductosPorCategoria() {
    return this.state.productos.filter((p) => p.categoria === this.state.categoria);
  }

  private maskCard(num?: string): string {
    const s = (num ?? "").replace(/\s+/g, "");
    if (!s) return "";
    return s.length >= 4 ? `**** **** **** ${s.slice(-4)}` : s;
  }

  private moneyToNumber(d?: Dinero | number | null): number | undefined {
    if (d == null) return undefined;
    if (typeof d === "number") return d;
    const n = Number(d.monto ?? 0);
    return isFinite(n) ? n : undefined;
  }

  private formatMoney(v: any) {
    const num = typeof v === "number" ? v : parseFloat(v ?? "0");
    if (!isFinite(num)) return "";
    try {
      return new Intl.NumberFormat("es-MX", { style: "currency", currency: "MXN", maximumFractionDigits: 2 }).format(num);
    } catch {
      return `${num}`;
    }
  }

  private safe(v: any) {
    if (v === null || v === undefined) return "";
    return String(v);
  }

  private showGlobalProgress(msg?: string) {
    try { (window as any).Xrm?.Utility?.showProgressIndicator?.(msg); } catch {}
  }
  private closeGlobalProgress() {
    try { (window as any).Xrm?.Utility?.closeProgressIndicator?.(); } catch {}
  }

  private asBool(v: any): boolean {
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v !== 0;
    const s = String(v ?? "").trim().toLowerCase();
    return s === "true" || s === "1" || s === "yes" || s === "si";
  }

  private cleanGuid(id: string | null | undefined): string {
    return (id ?? "").toString().trim().replace(/[{}]/g, "").toLowerCase();
  }

  private getFluentSelectedValue(el: any): string | null {
    try {
      const so = el?.selectedOptions;
      if (so && so.length) {
        const v = (so[0] as any).value ?? so[0]?.getAttribute?.("value");
        if (v !=null && v !== "") return String(v);
      }
      const v2 = el?.currentValue ?? el?.value;
      return v2 ? String(v2) : null;
    } catch {
      return null;
    }
  }

  // ======== Outputs ========
  public getOutputs(): IOutputs {
    return {
      jsonPersona: this.outJsonPersona,
      outMiddleName: this.outMiddleName,
      outLastName: this.outLastName,
      outFirstName: this.outFirstName,
      outEjecutivoTitular: this.outEjecutivoTitular,
      outAntiguedad: this.outAntiguedad,
      outEmail: this.outEmail,
      outMobile: this.outMobile,

      outUsuarioBancaElectronica: this.outUsuarioBancaElectronica,
      outTenenciaProductos: this.outTenenciaProductos,

      outGenderCode: this.outGenderCode,
      outMarcaDeVulnerabilidad: this.outMarcaDeVulnerabilidad,

      outSegmento: this.outSegmento,
      outSucursal: this.outSucursal,
      outSucursalAlta: this.outSucursalAlta,
      customerid: this.customerid,
    } as IOutputs;
  }

  public destroy(): void {}
}