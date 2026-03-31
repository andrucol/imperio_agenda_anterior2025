/**
 * ============================================
 * baseDatos.gs - CONFIGURACIÓN GLOBAL DEL SISTEMA
 * ============================================
 * 
 * Este archivo centraliza TODAS las constantes y variables
 * de configuración del proyecto. Cualquier cambio de ID,
 * horario o color se hace AQUÍ y se propaga automáticamente
 * a los demás archivos (codigo.gs, agendamientoWeb.gs, etc.).
 *
 * Fecha de creación: 30 de noviembre de 2025
 */

// ─────────────────────────────────────────────
// 1. BASE DE DATOS DE CLIENTES (Google Sheets externo)
// ─────────────────────────────────────────────
// ID del libro de Google Sheets que almacena la información de todos los clientes registrados.
const idBaseDeDatosClientes = "1g43H_UKFwuuWHqFLhxPGHWhGmyU0fCXwKBzDffdA2Ng";
// Nombre de la hoja (pestaña) dentro de ese libro donde están los datos de clientes.
const hojaBaseDeDatosClientes = "General";

// ─────────────────────────────────────────────
// 2. BASE DE DATOS DE AGENDA (Google Sheets externo)
// ─────────────────────────────────────────────
// ID del libro de Google Sheets que almacena todas las citas/agendamientos.
const idBaseDeDatosAgenda = "1_f4uTfUENXCGFdd0eUnu5jZQJTiJOC2dGh_HWuMz9Bs";
// Nombre de la hoja (pestaña) dentro de ese libro donde se guardan los reportes de citas.
const hojaBaseDeDatosAgenda = "Reporte";

// ─────────────────────────────────────────────
// 3. PARÁMETROS DE SEDE
// ─────────────────────────────────────────────
// Nombre de la sede que aparece en los mensajes de WhatsApp a los clientes.
const sedeMensaje = "OFICINA";

// ─────────────────────────────────────────────
// 4. HORARIO LABORAL ESTÁNDAR
// ─────────────────────────────────────────────
// Array con las franjas horarias que se despliegan en las agendas.
// Si se necesita agregar o quitar un horario (ej: abrir a las "7:30 AM"),
// solo se modifica este array y TODAS las agendas lo reflejarán al instante.
const HORAS_LABORALES = [
  "8 AM",
  "8:30 AM",
  "9 AM",
  "9:30 AM",
  "10 AM",
  "10:30 AM",
  "11 AM",
  "11:30 AM",
  "12 PM",
  "12:30 PM",
  "1 PM",
  "1:30 PM",
  "2 PM",
  "2:30 PM",
  "3 PM",
  "3:30 PM",
  "4 PM",
  "4:30 PM",
  "5 PM",
  "5:30 PM",
  "6 PM",
  "6:30 PM",
  "7 PM",
];

// ─────────────────────────────────────────────
// 5. COLORES PASTEL PARA VISUALIZACIÓN
// ─────────────────────────────────────────────
// Paleta de colores que se asigna cíclicamente a los profesionales
// en las vistas de día (Dia_1, Dia_2) y mes (Mes), para distinguir
// visualmente a cada profesional con un color diferente en las filas.
const COLORES_PASTEL_AGENDA = [
  "#FFFFFF", // Blanco (fila base)
  "#FADADD", // Rosa pálido
  "#E0F7FA", // Azul claro cielo
  "#FFFFE0", // Amarillo claro
  "#D7FFDB", // Verde menta
  "#F3E5F5", // Lavanda pálido
  "#FFF9C4", // Crema
  "#FFE0B2", // Durazno pálido
];
