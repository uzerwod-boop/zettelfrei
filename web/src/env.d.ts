// Kernmodule werden als klassische Skripte geladen (eine Quelle mit der klassischen Oberfläche)
declare const ZF: any, ZFE: any, ZFD: any, ZFX: any, AI: any, XLSX: any, ExcelJS: any, pdfjsLib: any;
interface Window { ZF_EXAMPLES?: Record<string, string>; ZF_DOCS?: Record<string, string>; claude?: any; ExcelJS?: any; XLSX?: any; }
declare module "*.css";
