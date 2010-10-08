VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm FrmAdmin 
   BackColor       =   &H00800000&
   Caption         =   "-"
   ClientHeight    =   870
   ClientLeft      =   60
   ClientTop       =   -255
   ClientWidth     =   9915
   Icon            =   "FrmAdmin.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "FrmAdmin.frx":27A2
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock wsServidor 
      Left            =   60
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   888
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Picture         =   "FrmAdmin.frx":17BF4
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "FrmAdmin.frx":18A46
            Object.ToolTipText     =   "Origen de datos...."
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDlMain 
      Left            =   720
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu AC0 
      Caption         =   "&Archivo"
      WindowList      =   -1  'True
      Begin VB.Menu AC01 
         Caption         =   "Seleccionar Inmueble"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu ACr 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu AC02 
         Caption         =   "Seleccionar Usuario"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC021 
         Caption         =   "Cambiar Contraseña"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC03 
         Caption         =   "Tablas del Sistema"
         Enabled         =   0   'False
         Begin VB.Menu AC0301 
            Caption         =   "Ciudades"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu AC0301 
            Caption         =   "Estados"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu AC0301 
            Caption         =   "Bancos"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu AC0301 
            Caption         =   "Ocupación o Cargos"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu AC0301 
            Caption         =   "Tipo de Inmueble"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu AC0301 
            Caption         =   "Actividad"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu AC0301 
            Caption         =   "Ramo"
            Enabled         =   0   'False
            Index           =   6
         End
         Begin VB.Menu AC0301 
            Caption         =   "Estado Civil"
            Enabled         =   0   'False
            Index           =   7
         End
         Begin VB.Menu AC0301 
            Caption         =   "Tipos de Contratos"
            Enabled         =   0   'False
            Index           =   8
         End
         Begin VB.Menu AC0301 
            Caption         =   "Condiciones de Pago"
            Enabled         =   0   'False
            Index           =   9
         End
      End
      Begin VB.Menu AC04 
         Caption         =   "Mensajería SAC"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC05 
         Caption         =   "Actualizar Deudas"
         Visible         =   0   'False
      End
      Begin VB.Menu ACr2 
         Caption         =   "-"
      End
      Begin VB.Menu AC07 
         Caption         =   "Pausar Sesión SAC..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu AC06 
         Caption         =   "Salir"
         Enabled         =   0   'False
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu AC1 
      Caption         =   "&Condominio"
      Begin VB.Menu AC101 
         Caption         =   "Ficha de Inmueble"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC102 
         Caption         =   "Ficha del Propietario"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC103 
         Caption         =   "Catálogo de Gastos"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC104 
         Caption         =   "Catálogo de Fondos"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC107 
         Caption         =   "Avisos de Cobro [Cuerpo del Mensaje]"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC105 
         Caption         =   "Editar Cartas de Morosidad"
         Enabled         =   0   'False
         Begin VB.Menu AC1050 
            Caption         =   "1.- Carta 3 meses"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu AC1050 
            Caption         =   "2.- Carta más 3 meses"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu AC1050 
            Caption         =   "3.- Telegramas"
            Enabled         =   0   'False
            Index           =   2
         End
      End
      Begin VB.Menu AC108 
         Caption         =   "Informe Económico"
      End
      Begin VB.Menu AC115 
         Caption         =   "Publicar Supervisión"
      End
      Begin VB.Menu AC109 
         Caption         =   "Enviar Email"
      End
      Begin VB.Menu AC114 
         Caption         =   "Enviar SMS"
      End
      Begin VB.Menu AC106 
         Caption         =   "Consultas y Reportes"
         Enabled         =   0   'False
         Begin VB.Menu AC10601 
            Caption         =   "Ficha de Inmuebles"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC10602 
            Caption         =   "Lista de Inmuebles"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC10603 
            Caption         =   "Ficha de Propietarios"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC10604 
            Caption         =   "Lista de Propietarios"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC10605 
            Caption         =   "Catálogo Concepto de Gastos"
            Enabled         =   0   'False
            Begin VB.Menu AC1060501 
               Caption         =   "Todos Los Conceptos"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu AC1060501 
               Caption         =   "Gastos Comunes"
               Enabled         =   0   'False
               Index           =   1
            End
            Begin VB.Menu AC1060501 
               Caption         =   "Gastos No Comunes"
               Enabled         =   0   'False
               Index           =   2
            End
            Begin VB.Menu AC1060501 
               Caption         =   "Gastos Fijos"
               Index           =   3
            End
            Begin VB.Menu AC1060501 
               Caption         =   "Cuentas de Fondos"
               Enabled         =   0   'False
               Index           =   4
            End
         End
         Begin VB.Menu AC10608 
            Caption         =   "Estado de Cuenta de Fondos"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC10609 
            Caption         =   "Estadísticas por Condominio"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC10610 
            Caption         =   "Reporte de Finiquito"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuEspace 
         Caption         =   "-"
      End
      Begin VB.Menu AC110 
         Caption         =   "Ver firma Junta de Condominio"
      End
      Begin VB.Menu AC111 
         Caption         =   "Diario..."
      End
      Begin VB.Menu AC112 
         Caption         =   "Busqueda Avanazada..."
      End
      Begin VB.Menu AC113 
         Caption         =   "Correspondencia"
      End
   End
   Begin VB.Menu AC8 
      Caption         =   "Al&quileres"
      Begin VB.Menu AC81 
         Caption         =   "Oferta"
         Index           =   0
      End
   End
   Begin VB.Menu AC2 
      Caption         =   "Cuentas por &Pagar"
      Begin VB.Menu AC201 
         Caption         =   "Ficha del Proveedor"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC202 
         Caption         =   "Recepción de Facturas"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC203 
         Caption         =   "Registrar Remesa"
         Begin VB.Menu AC2031 
            Caption         =   "Servicios"
            Index           =   0
         End
         Begin VB.Menu AC2031 
            Caption         =   "Gastos Menores"
            Index           =   1
            Begin VB.Menu AC20311 
               Caption         =   "Registrar"
               Index           =   0
            End
            Begin VB.Menu AC20311 
               Caption         =   "Procesar"
               Index           =   1
            End
            Begin VB.Menu AC20311 
               Caption         =   "Reporte General"
               Index           =   2
            End
         End
      End
      Begin VB.Menu AC204 
         Caption         =   "Asignación de Pagos"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC206 
         Caption         =   "Cronograma de Pagos"
         Begin VB.Menu AC2061 
            Caption         =   "Cheques en Tránsito"
            Index           =   0
         End
         Begin VB.Menu AC2061 
            Caption         =   "Cheques por Pagar"
            Index           =   1
         End
         Begin VB.Menu AC2061 
            Caption         =   "Cheques Pagados"
            Index           =   2
         End
         Begin VB.Menu AC2061 
            Caption         =   "Plan de Pagos"
            Index           =   3
         End
      End
      Begin VB.Menu AC207 
         Caption         =   "Agenda Telefónica"
         Enabled         =   0   'False
         Begin VB.Menu AC20701 
            Caption         =   "Proveedores"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu AC20701 
            Caption         =   "Propietarios"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu AC20701 
            Caption         =   "Junta de Condominio"
            Enabled         =   0   'False
            Index           =   2
         End
      End
      Begin VB.Menu AC209 
         Caption         =   "Emisión Orden de Pago"
      End
      Begin VB.Menu AC210 
         Caption         =   "[CxP] Consultas y Rep."
         Enabled         =   0   'False
         Begin VB.Menu AC21001 
            Caption         =   "Administrativos"
            Enabled         =   0   'False
            Begin VB.Menu AC2100101 
               Caption         =   "Lista de Proveedores"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100102 
               Caption         =   "Factura Pendiente Remesa"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100104 
               Caption         =   "Agenda Telefónica"
               Enabled         =   0   'False
               Begin VB.Menu AC210010401 
                  Caption         =   "Proveedores"
                  Enabled         =   0   'False
               End
               Begin VB.Menu AC210010402 
                  Caption         =   "Propietarios"
                  Enabled         =   0   'False
               End
               Begin VB.Menu AC210010403 
                  Caption         =   "Junta Condominio"
                  Enabled         =   0   'False
               End
            End
            Begin VB.Menu AC2100105 
               Caption         =   "Relación de Facturas Recibidas"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100106 
               Caption         =   "Relación de Cheques"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100107 
               Caption         =   "Honorarios Facturación"
               Begin VB.Menu AC21001070 
                  Caption         =   "Cuenta Pote"
                  Index           =   0
               End
               Begin VB.Menu AC21001070 
                  Caption         =   "Cuenta Separada"
                  Index           =   1
               End
               Begin VB.Menu AC21001070 
                  Caption         =   "General"
                  Index           =   2
               End
            End
         End
         Begin VB.Menu AC21002 
            Caption         =   "Operativos"
            Enabled         =   0   'False
            Begin VB.Menu AC2100202 
               Caption         =   "Reporte de Transacciones"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100203 
               Caption         =   "Estado de Cuenta por Proveedor"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100204 
               Caption         =   "Estado de Cuenta por Concepto"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100205 
               Caption         =   "Cuentas por Pagar"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100206 
               Caption         =   "Libro de Compras"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100207 
               Caption         =   "Resumen Anual de Gastos"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100208 
               Caption         =   "Análisis de Vencimientos"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100209 
               Caption         =   "Retenciones de I.S.L.R."
               Enabled         =   0   'False
            End
            Begin VB.Menu AC2100210 
               Caption         =   "Estadisticas de Compras"
               Enabled         =   0   'False
            End
         End
      End
      Begin VB.Menu AC211 
         Caption         =   "-"
      End
      Begin VB.Menu AC212 
         Caption         =   "Control Cheques Emitidos"
      End
   End
   Begin VB.Menu AC3 
      Caption         =   "&Facturación"
      Begin VB.Menu AC301 
         Caption         =   "Asignación de Gastos"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC313 
         Caption         =   "Novedades..."
         Enabled         =   0   'False
      End
      Begin VB.Menu AC314 
         Caption         =   "Revisión....."
         Enabled         =   0   'False
      End
      Begin VB.Menu AC303 
         Caption         =   "Prefacturación"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC304 
         Caption         =   "Parámetros de Facturación"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC305 
         Caption         =   "Registrar Gastos no Comunes"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC306 
         Caption         =   "Emisión de Facturas"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC307 
         Caption         =   "Revertir Facturación"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC309 
         Caption         =   "Cartas y Telegramas"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC310 
         Caption         =   "[Fact] Consultas y Reportes"
         Enabled         =   0   'False
         Begin VB.Menu AC31001 
            Caption         =   "Facturación"
            Enabled         =   0   'False
            Begin VB.Menu AC3100101 
               Caption         =   "Aviso de Cobro"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC3100108 
               Caption         =   "Aviso de Cobro (vía e-mail)"
            End
            Begin VB.Menu AC3100102 
               Caption         =   "Pre-Recibo"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC3100103 
               Caption         =   "Reporte de Facturación"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC3100104 
               Caption         =   "Análisis de Facturación"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC3100105 
               Caption         =   "Lista de Gastos No Comunes"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC3100107 
               Caption         =   "Control Facturación"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC3100106 
               Caption         =   "Paquete Completo"
               Enabled         =   0   'False
            End
            Begin VB.Menu AC3100109 
               Caption         =   "Libro de Ventas (I.V.A.)"
            End
         End
         Begin VB.Menu AC31003 
            Caption         =   "Resumen de Gastos Mensual"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC31004 
            Caption         =   "Reporte por Concepto de Gastos"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC31006 
            Caption         =   "Facturas (I.V.A.)"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC31007 
            Caption         =   "Estadísticas Mensuales"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu AC311 
         Caption         =   "-"
      End
      Begin VB.Menu AC312 
         Caption         =   "Cuadre Fondo-Deuda"
         Begin VB.Menu AC3121 
            Caption         =   "Imprimir Fondo-Deuda"
            Index           =   0
         End
         Begin VB.Menu AC3121 
            Caption         =   "Cuadre Deuda"
            Index           =   1
         End
         Begin VB.Menu AC3121 
            Caption         =   "Cuadre Fondo"
            Index           =   2
         End
      End
   End
   Begin VB.Menu AC4 
      Caption         =   "Ca&ja"
      Begin VB.Menu AC400 
         Caption         =   "Abrir Caja"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu AC400 
         Caption         =   "Cobranza por Caja [Bs.Fuertes]"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu AC400 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu AC400 
         Caption         =   "Cuadre de Caja"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu AC400 
         Caption         =   "Portadas de Caja"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu AC400 
         Caption         =   "[Caja] Consultas y Rep."
         Enabled         =   0   'False
         Index           =   5
         Begin VB.Menu AC4002 
            Caption         =   "&1.- Reporte General"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu AC4002 
            Caption         =   "&2.- Resumen de Caja"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu AC4002 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu AC4002 
            Caption         =   "Depósitos en Tránsito"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu AC4002 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu AC4002 
            Caption         =   "Emisión Canc. de Gastos"
            Enabled         =   0   'False
            Index           =   6
         End
         Begin VB.Menu AC4002 
            Caption         =   "Re-Impresión Canc. de Gastos"
            Enabled         =   0   'False
            Index           =   7
         End
      End
      Begin VB.Menu AC400 
         Caption         =   "Pagos WEB"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu AC400 
         Caption         =   "Autorizar..."
         Enabled         =   0   'False
         Index           =   7
         Begin VB.Menu AC4001 
            Caption         =   "Deducciones"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu AC4001 
            Caption         =   "Cierre de Caja"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu AC400 
         Caption         =   "Cerrar Caja"
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu AC400 
         Caption         =   "Aplicar Abonos..."
         Enabled         =   0   'False
         Index           =   9
      End
      Begin VB.Menu AC400 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu AC400 
         Caption         =   "Utilidades..."
         Enabled         =   0   'False
         Index           =   11
      End
      Begin VB.Menu AC400 
         Caption         =   "Cobranza por Caja [Bolívares]"
         Enabled         =   0   'False
         Index           =   12
         Visible         =   0   'False
      End
   End
   Begin VB.Menu ACCC 
      Caption         =   "Cuentas x C&obrar"
      Begin VB.Menu AC404 
         Caption         =   "Estado de Cuenta"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu AC404 
         Caption         =   "Avisos de Cobro"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu AC404 
         Caption         =   "Devolución de Cheque"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu AC404 
         Caption         =   "Consulta Administrativa"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu AC404 
         Caption         =   "Asignar Cobrador al Cliente"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu AC404 
         Caption         =   "Departamento Jurídico"
         Index           =   5
         Begin VB.Menu AC4041 
            Caption         =   "Seguimiento Gestión de Cobro"
            Index           =   0
         End
         Begin VB.Menu AC4041 
            Caption         =   "Registro de Pagos"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu AC4041 
            Caption         =   "Convenio de Pago"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu AC4041 
            Caption         =   "Emisión de Giros"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu AC4041 
            Caption         =   "Abogados..."
            Index           =   4
         End
         Begin VB.Menu AC4041 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu AC4041 
            Caption         =   "[Convenio] Consultas y Reportes"
            Index           =   6
         End
      End
      Begin VB.Menu AC404 
         Caption         =   "[CxC] Consultas y Reportes"
         Enabled         =   0   'False
         Index           =   6
         Begin VB.Menu AC4042 
            Caption         =   "Administrativos"
            Enabled         =   0   'False
            Index           =   0
            Begin VB.Menu AC40421 
               Caption         =   "Relación de CxC para Cobrador"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu AC40421 
               Caption         =   "Relación Recibos por Enviar"
               Enabled         =   0   'False
               Index           =   1
            End
            Begin VB.Menu AC40421 
               Caption         =   "Estado de Cuenta por Inmueble"
               Enabled         =   0   'False
               Index           =   2
            End
            Begin VB.Menu AC40421 
               Caption         =   "Reporte de Atrasos"
               Enabled         =   0   'False
               Index           =   3
            End
            Begin VB.Menu AC40421 
               Caption         =   "Reporte de Honorarios (Legal)"
               Enabled         =   0   'False
               Index           =   4
            End
            Begin VB.Menu AC40421 
               Caption         =   "Consulta de Honorarios"
               Enabled         =   0   'False
               Index           =   5
            End
            Begin VB.Menu AC40421 
               Caption         =   "Reporte de Morosos"
               Enabled         =   0   'False
               Index           =   6
            End
            Begin VB.Menu AC40421 
               Caption         =   "Análisis de Vencimiento"
               Index           =   7
            End
            Begin VB.Menu AC40421 
               Caption         =   "Intereses Descontados"
               Index           =   8
            End
         End
         Begin VB.Menu AC4042 
            Caption         =   "Estadisticos"
            Enabled         =   0   'False
            Index           =   1
            Begin VB.Menu AC40422 
               Caption         =   "Estado Financiero Anual"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu AC40422 
               Caption         =   "Reporte de Cobranza Efectiva"
               Enabled         =   0   'False
               Index           =   1
            End
            Begin VB.Menu AC40422 
               Caption         =   "Reporte de Deuda Mensual"
               Enabled         =   0   'False
               Index           =   2
            End
            Begin VB.Menu AC40422 
               Caption         =   "Resumen Anual de Cobros"
               Enabled         =   0   'False
               Index           =   3
            End
            Begin VB.Menu AC40422 
               Caption         =   "Relación  Fondo-Deuda"
               Enabled         =   0   'False
               Index           =   4
            End
            Begin VB.Menu AC40422 
               Caption         =   "Estadisticas Mensuales de Deuda"
               Enabled         =   0   'False
               Index           =   5
            End
         End
      End
      Begin VB.Menu AC404 
         Caption         =   "Aviso Legal"
         Index           =   7
      End
   End
   Begin VB.Menu AC5 
      Caption         =   "&Bancos"
      Begin VB.Menu AC501 
         Caption         =   "Ficha Bancaria"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC502 
         Caption         =   "Cuentas Bancarias"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC503 
         Caption         =   "Libro Diario"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC504 
         Caption         =   "Administrador de Chequeras"
         Enabled         =   0   'False
         Begin VB.Menu AC50401 
            Caption         =   "Registro de Chequeras"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC50402 
            Caption         =   "Asignación de Chequeras"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu AC505 
         Caption         =   "Registro Cheques Devueltos"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC506 
         Caption         =   "Conciliación Bancaria"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC507 
         Caption         =   "[Banco] Consultas y Reportes"
         Enabled         =   0   'False
         Begin VB.Menu AC50701 
            Caption         =   "Lista de Bancos"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC50702 
            Caption         =   "Consulta de Saldos"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC50703 
            Caption         =   "Reporte de Transacciones"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC50704 
            Caption         =   "Estados de Cuentas"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC50705 
            Caption         =   "Disponibilidad Bancaria"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC50706 
            Caption         =   "Relación de Cheques Devueltos"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC50707 
            Caption         =   "Relación de Chequeras"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu AC508 
         Caption         =   "-"
      End
      Begin VB.Menu AC509 
         Caption         =   "Buscador de Depósitos"
      End
   End
   Begin VB.Menu AC6 
      Caption         =   "&Nómina"
      Begin VB.Menu AC601 
         Caption         =   "Ficha del Trabajador"
         Index           =   0
      End
      Begin VB.Menu AC601 
         Caption         =   "Nómina"
         Index           =   1
         Begin VB.Menu AC6014 
            Caption         =   "Pre-Nómina"
            Index           =   0
         End
         Begin VB.Menu AC6014 
            Caption         =   "Procesar Nómina"
            Index           =   1
         End
         Begin VB.Menu AC6014 
            Caption         =   "Novedades...."
            Index           =   2
         End
         Begin VB.Menu AC6014 
            Caption         =   "Cuenta inmueble"
            Index           =   3
         End
         Begin VB.Menu AC6014 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu AC6014 
            Caption         =   "Parámetros...."
            Index           =   5
         End
      End
      Begin VB.Menu AC601 
         Caption         =   "Prestaciones Sociales"
         Index           =   2
      End
      Begin VB.Menu AC601 
         Caption         =   "Vacaciones"
         Index           =   3
      End
      Begin VB.Menu AC601 
         Caption         =   "Aguinaldos"
         Index           =   4
      End
      Begin VB.Menu AC601 
         Caption         =   "Editar..."
         Index           =   5
         Begin VB.Menu AC6011 
            Caption         =   "Cargos"
            Index           =   0
         End
         Begin VB.Menu AC6011 
            Caption         =   "Contratos"
            Index           =   1
         End
      End
      Begin VB.Menu AC601 
         Caption         =   "Consultas y Reportes"
         Index           =   6
         Begin VB.Menu AC60601 
            Caption         =   "Lista de Empleados"
         End
         Begin VB.Menu AC60602 
            Caption         =   "Constancia de Trabajo"
         End
         Begin VB.Menu AC60603 
            Caption         =   "Nómina"
         End
         Begin VB.Menu AC60604 
            Caption         =   "Reporte por Conceptos"
         End
         Begin VB.Menu AC60607 
            Caption         =   "Reportes Bancarios"
            Begin VB.Menu AC6060701 
               Caption         =   "Transferencias"
            End
            Begin VB.Menu AC6060702 
               Caption         =   "Depósito de L.P.H."
            End
            Begin VB.Menu AC6060703 
               Caption         =   "S.S.O. y Paro Forzoso"
            End
         End
         Begin VB.Menu AC60608 
            Caption         =   "Fondo de Prestaciones Sociales"
         End
         Begin VB.Menu AC60609 
            Caption         =   "Reporte de Aquinaldos"
         End
      End
      Begin VB.Menu AC601 
         Caption         =   "Seguro Social"
         Index           =   7
         Begin VB.Menu AC6016 
            Caption         =   "Remesa"
            Index           =   0
         End
         Begin VB.Menu AC6016 
            Caption         =   "Nº de Empresa"
            Index           =   1
         End
         Begin VB.Menu AC6016 
            Caption         =   "Parámetros"
            Index           =   2
         End
      End
   End
   Begin VB.Menu AC7 
      Caption         =   "&Utilidades"
      Begin VB.Menu AC701 
         Caption         =   "Perfiles de Acceso"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC702 
         Caption         =   "Usuarios y Perfiles"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC703 
         Caption         =   "Datos de la Empresa"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC704 
         Caption         =   "Actualizar SAC Portátil"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC706 
         Caption         =   "Bitácora del Sistema"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC708 
         Caption         =   "Control de Asistencia"
         Enabled         =   0   'False
      End
      Begin VB.Menu AC705 
         Caption         =   "Parámetros del Sistema"
         Enabled         =   0   'False
         Begin VB.Menu AC70501 
            Caption         =   "Administrativos"
            Enabled         =   0   'False
         End
         Begin VB.Menu AC70502 
            Caption         =   "Operativos"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu AC707 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu AC707 
         Caption         =   "Quórum"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu AC707 
         Caption         =   "Modo Local"
         Index           =   2
      End
   End
   Begin VB.Menu Mante 
      Caption         =   "Mantenimiento"
      Visible         =   0   'False
      Begin VB.Menu New 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu Del 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu mnuBit 
      Caption         =   "Bitacora"
      Visible         =   0   'False
      Begin VB.Menu mnuBitacora 
         Caption         =   "Copiar"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuBitacora 
         Caption         =   "Resaltar"
         Index           =   1
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Imprimir"
      Visible         =   0   'False
      Begin VB.Menu mnuRelacion 
         Caption         =   "Relación Deuda"
      End
      Begin VB.Menu mnuConvenio 
         Caption         =   "Convenio"
      End
   End
End
Attribute VB_Name = "FrmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    '05/08/2002 Rev.-------------------------------------------------Módulo Formulario Principal
    Dim strRecibo As String           ' Pública a nivel de módulo guarda Id de Trans. de caja
    Dim CnnCaja As ADODB.Connection       ' Conexión Pública a nivel de módulo
    Dim rstCaja As ADODB.Recordset             ' Recorset público a nivel de módulo
    Dim lngINI As Long, lngFIN As Long
    Const lABONO_INICIAL% = 0
    Const lABONO_FINAL% = 1
    Dim af As String        'Código de cuenta abono a futuro
    Dim Mensaje As String
    Dim ctlReport As Object
    Dim NewUser As Boolean
    Public objRst As New ADODB.Recordset
    Public ObjRstNom As New ADODB.Recordset
    Private WithEvents poSendMail As clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
    Private bSendFailed     As Boolean
    
    '
    '------------------------------------------------------------
    Private Sub AC01_Click()    'Seleccionar Inmueble
    '------------------------------------------------------------
    'variables locales
    Dim Frm As Form
    '
    For Each Frm In Forms
    '
        If Frm.Name <> "FrmSelCon" And Frm.Name <> "FrmAdmin" Then
            If Frm.Tag = "1" Then Unload Frm
        End If
        '
    Next
    Call Muestra_Formulario(FrmSelCon, "Click Selección Inmueble")
    '
    End Sub

'    '----------------------------------------------------------------------------
'    Private Sub AC010506_Click()    'Print Rel. Deuda / Propietario
'    '----------------------------------------------------------------------------
'    '
'    mcTitulo = "Relación de Deuda por Propietario Inm:" & gcCodInm
'    mcReport = "LisDeuProp.Rpt"
'    mcOrdCod = "+{Propietarios.Codigo}"
'    mcOrdAlfa = "+{Propietarios.Nombre}"
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    '
'    End Sub

'    '---------------------------------------------------------------------------------
'    Private Sub AC010507_Click()    'Print Rel. Deuda / Fondo Reserva
'    '--------------------------------------------------------------------------------
'    '
'    mcTitulo = "Relación entre Deuda y Fondo"
'    mcReport = "LisDeuFon.Rpt"
'    mcOrdCod = "+{Inmueble.CodInm}"
'    mcOrdAlfa = "+{Inmueble.Nombre}"
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir" & mcTitulo)
'    '
'    End Sub

'    'Revisar este código
'    Private Sub AC010607_Click()
'    mcTitulo = "Relación entre Fondo y Deuda"
'    mcReport = "Lisdeufon.Rpt"
'    mcOrdCod = ""
'    mcOrdAlfa = ""
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

    Private Sub AC02_Click()
    'variables locales
    NewUser = True
    Unload FrmAdmin
    NewUser = False
    Load frmAcceso
    frmAcceso.fraAcceso(0).Visible = True
    frmAcceso.fraAcceso(3).Visible = True
    'Call Muestra_Formulario(frmAcceso, "Seleccionar Nuevo Usuario...")
    '
    End Sub

    Private Sub AC1004_Click(): Call Muestra_Formulario(FrmEditaCartas, "Click Edición de Cartas")
    End Sub

    Private Sub AC021_Click(): Call Muestra_Formulario(frmPass, "Click Cambio de Contraseña")
    End Sub

    
    Private Sub AC0301_Click(Index As Integer)
    'Matriz de Menus Tablas del Sistema
    Select Case Index
    '
        Case 0  'Ciudades
        '-------------------------
            mcTitulo = "Tabla Ciudades"
            mcTablas = "Ciudades"
        Case 1  'Estados
        '-------------------------
            mcTitulo = "Tabla Estados"
            mcTablas = "Estados"
        Case 2  'Bancos
        '-------------------------
            mcTitulo = "Tabla Bancos"
            mcTablas = "Bancos"
        Case 3  'Ocupación o cargos
        '-------------------------
            mcTitulo = "Tabla Ocupación o Cargos"
            mcTablas = "Cargos"
        Case 4  'Tipo Inmueble
        '-------------------------
            mcTitulo = "Tabla Tipo de Inmueble"
            mcTablas = "TipoInm"
        Case 5  'actividad
        '-------------------------
            mcTitulo = "Tabla Actividades"
            mcTablas = "Actividad"
        Case 6  'Ramo
        '-------------------------
            mcTitulo = "Tabla Ramo"
            mcTablas = "Ramo"
        Case 7  'Estado Civil
        '-------------------------
            mcTitulo = "Tabla Estado Civil"
            mcTablas = "Civil"
        Case 8  'Tipo Contratos
        '-------------------------
            mcTitulo = "Tabla Tipo de Contratos"
            mcTablas = "Contratos"
        Case 9 'Condiciones de Pago
        '-------------------------
            mcTitulo = "Tabla Condiciones de Pago"
            mcTablas = "Condicion"
    End Select
    Call Muestra_Formulario(FrmTabla, "Click " & mcTitulo)
    '
    End Sub

    Private Sub AC04_Click()
    'messenger de sac
    Dim Frm As New frmMsg
    '
    Frm.cmd(1).Enabled = True
    Frm.Timer1(1).Interval = 10
    Frm.Show vbModeless, FrmAdmin
    '
    End Sub

    Private Sub AC07_Click()
    'pausar sac
    Call Muestra_Formulario(frmPausa, "Pausar Sac.......")
    End Sub

    Private Sub AC101_Click(): Call Muestra_Formulario(FrmInmueble, "Click Ficha Inmueble")
    End Sub

    Private Sub AC102_Click()
    If Not Estado Then Call Muestra_Formulario(FrmPropietario, "Click Ficha Propietarios del " _
    & gcCodInm)
    End Sub

    Private Sub AC06_Click()    'salir
    'variables locales
    Dim Frm As Form
    '
    For Each Frm In Forms
        If Frm.Name <> "FrmAdmin" Then Unload Frm
    Next
    Unload Me
    End
    '
    End Sub

    Private Sub AC103_Click()
    If Not Estado Then Call Muestra_Formulario(FrmTgasto, "Click Catálogo de Gastos Inm.: " _
    & gcCodInm)
    End Sub

    Private Sub AC104_Click()
    If Not Estado Then Call Muestra_Formulario(FrmTfondos, "Click Catálogo de Fondos Inm.: " _
    & gcCodInm)
    End Sub

'    Private Sub AC10501_Click()
'    mcTitulo = "Ficha de Inmuebles"
'    mcReport = "FichaInm.Rpt"
'    mcOrdCod = "+{Inmueble.CodInm}"
'    mcOrdAlfa = "+{Inmueble.Nombre}"
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

'    Private Sub AC10502_Click()
'    mcTitulo = "Lista de Inmuebles"
'    mcReport = "ListaInm.Rpt"
'    mcOrdCod = "+{Inmueble.CodInm}"
'    mcOrdAlfa = "+{Inmueble.Nombre}"
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub


'    Private Sub AC105051_Click()
'    mcTitulo = "Catálogo de Conceptos de Gastos Inm:" & gcCodInm
'    mcReport = "ListaGas.Rpt"
'    mcOrdCod = "+{Tgastos.CodGasto}"
'    mcOrdAlfa = "+{Tgastos.Titulo}"
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir  " & mcTitulo)
'    End Sub

'    Private Sub AC105052_Click()
'    mcTitulo = "Catálogo de Conceptos Comúnes Inm:" & gcCodInm
'    mcReport = "ListaGas.Rpt"
'    mcOrdCod = "+{Tgastos.CodGasto}"
'    mcOrdAlfa = "+{Tgastos.Titulo}"
'    mcCrit = "{Tgastos.Comun}"
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

'    Private Sub AC105053_Click()
'    mcTitulo = "Catálogo de Conceptos No Comúnes Inm:" & gcCodInm
'    mcReport = "ListaGas.Rpt"
'    mcOrdCod = "+{Tgastos.CodGasto}"
'    mcOrdAlfa = "+{Tgastos.Titulo}"
'    mcCrit = "not {Tgastos.Comun}"
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

'    Private Sub AC105054_Click()
'    mcTitulo = "Catálogo de Conceptos Fijos Inm:" & gcCodInm
'    mcReport = "ListaGas.Rpt"
'    mcOrdCod = "+{Tgastos.CodGasto}"
'    mcOrdAlfa = "+{Tgastos.Titulo}"
'    mcCrit = "{Tgastos.Fijo}"
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

'    Private Sub AC105055_Click()
'    mcTitulo = "Catálogo de Cuentas de Fondos Inm:" & gcCodInm
'    mcReport = "ListaGas.Rpt"
'    mcOrdCod = "+{Tgastos.CodGasto}"
'    mcOrdAlfa = "+{Tgastos.Titulo}"
'    mcCrit = "{Tgastos.Fondo}"
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

    Private Sub AC1050_Click(Index As Integer)
    'Matriz de Menus [Editar Cartas de Morosidad
    Select Case Index
    '
        Case 0  'Carta de 3 meses
        '--------------------
            Shell "WINWORD.EXE " & Left(gcPath, 17) & "\Docs\aviso1.doc", vbMaximizedFocus
            Call rtnBitacora("Editar Carta de Morosidad 3 Meses..")
            
        Case 1  'Carta más de 3 meses
        '--------------------
            Shell "WINWORD.EXE " & Left(gcPath, 7) & "Docs\aviso1.doc", vbMaximizedFocus
            Call rtnBitacora("Editar Carta de Morosidad más de 3 Meses..")
            
        Case 2  'Telegramas
        '--------------------
            Call Muestra_Formulario(FrmEditaCartas, "Editar Telegramas..")
            
    End Select
    '
    End Sub

    Private Sub AC10601_Click()
    mcTitulo = "Ficha de Inmuebles"
    mcReport = "FichaInm.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    End Sub

    Private Sub AC10602_Click()
    mcTitulo = "Lista de Inmuebles"
    mcReport = "ListaInm.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    End Sub

    Private Sub AC10603_Click()
    mcTitulo = "Ficha de Propietarios Inm:" & gcCodInm
    mcReport = "Fichapro.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir  " & mcTitulo)
    End Sub

    Private Sub AC10604_Click()
    mcTitulo = "Lista de Propietarios Inm:" & gcCodInm
    mcReport = "listapro.Rpt"
    mcOrdCod = "+{Propietarios.Codigo}"
    mcOrdAlfa = "+{Propietarios.Nombre}"
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    End Sub


    Private Sub AC1060501_Click(Index As Integer)   'Matriz de menús catálogo de gastos
    Dim strTitulo As String 'Variables locales
    If Estado Then Exit Sub
    Select Case Index
        Case 0  'Todos los gastos
        '--------------------
            mcCrit = ""
        Case 1  'Gastos Comunes
        '--------------------
            mcCrit = "{Tgastos.Comun}"
            strTitulo = "Gastos Comunes"
            
        Case 2  'Gastos No comunes
        '--------------------
            mcCrit = "Not {Tgastos.Comun}"
            strTitulo = "Gastos No Comunes"
            
        Case 3  'Gastos Fijos
        '--------------------
            mcCrit = "{Tgastos.Fijo}"
            strTitulo = "Gastos Fijos"
            
        Case 4  'Cuentas de fondos
        '--------------------
            mcCrit = "{Tgastos.Fondo}"
            strTitulo = "de Fondo"
            
    End Select
    '
    mcTitulo = "Catálogo de Cuentas " & strTitulo
    mcReport = "ListaGas.Rpt"
    mcOrdCod = "+{Tgastos.CodGasto}"
    mcOrdAlfa = "+{Tgastos.Titulo}"
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    '
    End Sub

    Private Sub AC10608_Click()
    If Estado Then Exit Sub
    Call Muestra_Formulario(FrmTfondos, "Click Edo.Cta.Fondos")
    FrmTfondos.sstFondo.tab = 1
    End Sub

    Private Sub AC107_Click()
    Call Muestra_Formulario(frmEmail, "Avisos de Cobro [Cuerpo del Mensaje]")
    End Sub
    
    Private Sub ac108_Click()   'informe económico
    If Not Estado Then Call Muestra_Formulario(frmIEco, "Click Informe Econónimco Inm:" & gcCodInm)
    End Sub

    Private Sub AC109_Click()
    'enviar email
    Call Muestra_Formulario(frmMail, "Click enviar email")
    End Sub

    Private Sub AC110_Click()   'ver la firma de la junta de condominio
    'varibales locales
    Dim Formulario As New frmFirma
    '
    If Dir(gcPath & gcUbica & "\reportes\" & gcCodInm & ".gif") = "" Then
        MsgBox "No existe información asociada con este inmueble", vbInformation, App.ProductName
    Else
        Formulario.strFichero = gcPath & gcUbica & "\reportes\" & gcCodInm & ".gif"
        Call Muestra_Formulario(Formulario, "Ver firma Junta Condominio " & gcCodInm)
    End If
    '
    End Sub

    Private Sub AC111_Click()
    Load frmDiario
    End Sub

Private Sub AC112_Click()
'buscar propietarios
Call Muestra_Formulario(frmBAvan, "Busqueda Avanzada Propietario")
End Sub

    Private Sub AC113_Click()
    Call Muestra_Formulario(frmCorrespondencia, "Click Correspondencia")
    End Sub

    Private Sub AC114_Click()
    'enviar sms
    Call Muestra_Formulario(frmSMS, "Click enviar sms")
    End Sub

    Private Sub AC115_Click()
    '   publicar el informe de supervision en la pag. web
    Dim Frm As frmSupervision
    Set Frm = New frmSupervision
    Frm.Show vbModal, FrmAdmin
    End Sub

    Private Sub AC201_Click(): Call Muestra_Formulario(FrmProveedor, "Click Ficha Proveedores")
    End Sub

    Private Sub AC202_Click(): Call Muestra_Formulario(FrmFactura, "Click Recepción de Facturas")
    End Sub

    
Private Sub AC2031_Click(Index As Integer)
Select Case Index
    Case 0  'remesa de servicios
        Call Muestra_Formulario(FrmRemesa, "Click Registrar Remesa Servicios")
    'Case 1  'remesa de gastos fijos menores
        
End Select

End Sub

Private Sub AC20311_Click(Index As Integer)
Select Case Index
    Case 0  'registrar
        Call Muestra_Formulario(frmGFMen, "Click Registrar Remesa Gastos Menores")
    Case 1  'procesar
        Call Procesar_GastosMenores
    Case 2  'reporte general
        Call reporte_gm_general
        
End Select
End Sub

'    Private Sub AC2051_Click(): Call Muestra_Formulario(FrmAgenda1, "Click Agenda Telefónica")
'    End Sub

'    Private Sub AC2052_Click()
'    Call Muestra_Formulario(FrmAgenda, "Click Agenda Telefónica Propietarios Inm.: " & gcCodInm)
'    End Sub

'    Private Sub AC2053_Click()
'    Call Muestra_Formulario(FrmAgenda2, "Click Agenda Telf. Junta de Cond. Inm.: " & gcCodInm)
'    End Sub

'    Private Sub AC211101_Click()
'    mcTitulo = "Listado de Proveedores"
'    mcReport = "LisProv.Rpt"
'    mcOrdCod = "+{Proveedores.Codigo}"
'    mcOrdAlfa = "+{Proveedores.Nombre}"
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

'    Private Sub AC211102_Click()
'    mcTitulo = "Reporte de Remesas Registradas"
'    mcReport = "LisRemReg.Rpt"
'    mcOrdCod = ""
'    mcOrdAlfa = ""
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

'    Private Sub AC211105_Click()
'    mcTitulo = "Relación de Facturas Recibidas"
'    mcReport = "LisRecFac.Rpt"
'    mcOrdCod = ""
'    mcOrdAlfa = ""
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

'    Private Sub AC212_Click(): Call Muestra_Formulario(FrmRemesa, "Cilck Registro de Remesas")
'    End Sub

'    Private Sub AC3003_Click()
'        Call Muestra_Formulario(FrmParamFact, "Parámetros de Fact. Inm.: " & gcCodInm)
'    End Sub

    Private Sub AC204_Click()
    '
    If Not Estado Then Call Muestra_Formulario(FrmAsignaPago, "Asignación de Pagos Inm.:" _
    & gcCodInm)
    '
    End Sub
    
    Private Sub AC2061_Click(Index%)
    ' variables locales
    Dim Formulario As frmCronoPago
    '
    Set Formulario = New frmCronoPago
    Select Case Index
    '
        Case 0  'cheques en tránsito
        '----------
            Formulario.Opcion = Index
            Call Muestra_Formulario(Formulario, "Cheques en Tránsito")
                
        Case 1  'cheques por entregar
        '----------
            Formulario.Opcion = Index
            Call Muestra_Formulario(Formulario, "Cheques por Entregar")
            
        Case 2  'Cheques Pagados
        '----------
            Formulario.Opcion = Index
            Call Muestra_Formulario(Formulario, "Cheques Pagados")
        
        Case 3  'plan de pagos
        '----------
            Call Muestra_Formulario(frmPlanPago, "Plan de pagos")
            '
        End Select
        '
    End Sub

    Private Sub AC20701_Click(Index As Integer) 'Matríz de Menús Agenda Telefónica
    '
    Select Case Index
    
        Case 0  'Proveedores
        '--------------------
            mcTitulo = "Agenda de Proveedores"
            mcReport = "Agenda.Prov.Rpt"
            mcOrdCod = ""
            mcOrdAlfa = ""
            mcCrit = ""
            Call Muestra_Formulario(FrmAgenda1, "Click Imprimir " & mcTitulo)
            
            Case 1  'Propietarios
        '--------------------
            If Not Estado Then Call Muestra_Formulario(FrmAgenda, "Click Agenda Propietarios In" _
            & "m:" & gcCodInm)
            
        Case 2  'Junta de Condominio
        '--------------------
            If Not Estado Then Call Muestra_Formulario(FrmAgenda2, "Click Agenda Junta de Condo" _
            & "minio Inm:" & gcCodInm)
        '
    End Select
    '
    End Sub

    Private Sub AC209_Click()
    'variables locales
    frmOrdPag.Show vbModeless, FrmAdmin
    End Sub

    Private Sub AC2100101_Click()
    'variables locales
    mcTitulo = "Lista de Proveedores"
    mcReport = "lisprov.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    
    End Sub
    
    Private Sub AC2100102_Click()
    'emisión de facturas no recibidas en remesa
    Call Muestra_Formulario(frmRemFal, "Facturas Faltantes Remesa")
    End Sub
    
    Private Sub AC210010401_Click()
    'variables locales
    mcTitulo = "Agenda Proveedores"
    mcReport = "agendaprov.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    End Sub

    Private Sub AC210010402_Click()
    'variables locales
    If Estado Then Exit Sub
    mcTitulo = "Agenda Propietarios"
    mcReport = "agendaprop.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    End Sub

    Private Sub AC210010403_Click()
    If Estado Then Exit Sub
    mcTitulo = "Agenda Junta Condominio"
    mcReport = "agendajuntacon.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    End Sub

    Private Sub AC2100105_Click()
    FrmReport.Frame1.Visible = True
    FrmReport.MskDesde = Date
    FrmReport.MskHasta = Date
    mcTitulo = "Reporte de Recepción de Facturas"
    mcReport = "RecepFc.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    End Sub

    Private Sub AC2100203_Click()   'estado de cuenta por proveedor
    frmProvEC.Show
    End Sub
    Private Sub AC2100204_Click()   'Estado de cuenta por código de gasto
    If Not Estado Then Call Muestra_Formulario(frmEdoCta, "Click Movimiento Cuenta")
    End Sub
    
    Private Sub AC212_Click()
    If gcNivel > nuAdministrador Then
        MsgBox "No cuenta con el perfíl para ver esta opción", vbCritical, "Acceso Denegado"
    Else
        Call Muestra_Formulario(frmchequecontrol, "Control Cheques Emitidos")
    End If
    
    End Sub

    Private Sub AC301_Click()
    If Not Estado Then Call Muestra_Formulario(FrmAsignaGasto, "Asignación de Gastos Inm.:" _
    & gcCodInm)
    End Sub

    Private Sub AC303_Click()
    FrmPeriodo.Show vbModeless, FrmAdmin
    Call rtnBitacora("Pre-Facturación Inm.: " & gcCodInm)
    'If Not Estado Then Call Muestra_Formulario(FrmPeriodo, "Pre-Facturación Inm.: " _
    & gcCodInm)
    End Sub

    Private Sub AC304_Click()
    Call Muestra_Formulario(FrmParamFact, "Parámetros de Facturación Inm.:" & gcCodInm)
    End Sub

'    '------------------------------
'    Private Sub AC308012_Click() '-
'    '------------------------------
'    '
'    If Estado Then Exit Sub
'    mcTitulo = "Resumen de Pre-Factura"
'    mcReport = "PreRecibo.Rpt"
'    mcOrdCod = ""
'    mcOrdAlfa = ""
'    mcCrit = ""
'    Call Muestra_Formulario(FrmPeriodo, "Impresión " & mcTitulo)
'    '
'    End Sub


    Private Sub AC305_Click()
    If Not Estado Then Call Muestra_Formulario(FrmNoComunes, "Cargar Gasto No Comun " _
    & "Inm.:" & gcCodInm)
    '
    End Sub

    Private Sub AC306_Click()
    If Not Estado Then Call Muestra_Formulario(FrmEmisionFactura, "Click Emisión Factura Inm.: " _
    & gcCodInm): strLlamada = "F"
    End Sub
    'Revertir Facturación
    Private Sub AC307_Click()
    If Estado Then Exit Sub
    strLlamada = "R"
    mcTitulo = "Revertir Facturación"
    Call Muestra_Formulario(FrmRepFact, "Revertir Facturación  Inm.: " & gcCodInm)
    End Sub

    Private Sub AC309_Click()
    Call Muestra_Formulario(FrmAvisos, "Click Notificación de Cobro")
    End Sub

    Private Sub AC3100101_Click()
    If Estado Then Exit Sub
    mcTitulo = "Aviso de Cobro"
    mcReport = "fact_Aviso.rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    strLlamada = "F"
    'FrmRepFact.eMail = False
    Call Muestra_Formulario(FrmRepFact, "Click Imprimir " & mcTitulo & " Inm.: " & gcCodInm)
    '
    End Sub

    '-------------------------------
    Private Sub AC3100102_Click() '-
    '-------------------------------
    '
    If Estado Then Exit Sub
    mcTitulo = "Pre-Recibo de Facturación"
    mcReport = "PreRecibo.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    strLlamada = "F"
    Call Muestra_Formulario(FrmRepFact, "Click Imprimir " & mcTitulo & " Inm.: " & gcCodInm)
    '
    End Sub

    Private Sub AC3100103_Click()
    '
    If Estado Then Exit Sub
    mcTitulo = "Reporte de Facturación"
    mcReport = "fact_mes.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    strLlamada = "F"
    Call Muestra_Formulario(FrmRepFact, "Click Imprimir " & mcTitulo & " Inm.: " & gcCodInm)
    '
    End Sub

    Private Sub AC3100104_Click()
    If Estado Then Exit Sub
    mcTitulo = "Análisis de Facturación"
    mcReport = "fact_analisis.rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    strLlamada = "F"
    Call Muestra_Formulario(FrmRepFact, "Click Imprimir " & mcTitulo & " Inm.: " & gcCodInm)
    End Sub
    
    Private Sub AC3100105_Click()
    If Estado Then Exit Sub
    mcTitulo = "Gastos No comunes"
    mcReport = "fact_GNC.rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    strLlamada = "F"
    Call Muestra_Formulario(FrmRepFact, "Click Imprimir " & mcTitulo & " Inm.: " & gcCodInm)
    End Sub
    
    Private Sub AC3100107_click()
    If Estado Then Exit Sub
    mcTitulo = "Control de Facturación"
    mcReport = "Fact_control.rpt"
    strLlamada = "F"
    Call Muestra_Formulario(FrmRepFact, "Click Imprimir " & mcTitulo & " Inm.: " & gcCodInm)
    End Sub
    
    Private Sub AC3100108_click()   'Aviso de cobro vía e-mail
    If Estado Then Exit Sub
    mcTitulo = "Aviso de Cobro"
    mcReport = "aviso_email"
    'strLlamada = "F"
    Call Muestra_Formulario(frmSelecInm, "Click Enviar " & mcTitulo & " vía email Inm.: " & gcCodInm)
    End Sub
    
    Private Sub AC3100106_Click()
    If Estado Then Exit Sub
    mcTitulo = "Paquete Completo Inm.: " & gcCodInm
    mcReport = "PAQUETECOMPLETO"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    strLlamada = "F"
    Call Muestra_Formulario(FrmRepFact, "Click Imprimir " & mcTitulo)
    End Sub
    
    Private Sub AC3100109_Click()
    mcTitulo = "Listado IVA"
    mcReport = "IVA"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    strLlamada = "F"
    Call Muestra_Formulario(FrmRepFact, "Click Imprimir " & mcTitulo)
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Genera un reporte de los gastos asignados por cheque al mes siguiente
    '   a facturar
    '---------------------------------------------------------------------------------------------
    Public Sub AC31003_Click()
        mcReport = "GM"
        strLlamada = "F"
        mcTitulo = "Reporte Gastos Mesuales"
        Call Muestra_Formulario(FrmRepFact, "Gastos Mensuales  Inm.: " & gcCodInm)
    End Sub

    Private Sub AC31006_Click()
    'impresion facturas IVA
    'Call Muestra_Formulario(frmFactIva, "Click Emisión Facturas IVA")
    frmFactIva.Show vbModeless, FrmAdmin
    Call rtnBitacora("Click Emisión Facturas IVA")
    
    End Sub


'    Private Sub AC31002_Click()
'    If Estado Then Exit Sub
'    mcTitulo = "Lista de Gastos No Comunes Inm.: " & gcCodInm
'    mcReport = "LisNoComun.Rpt"
'    mcDatos = gcPath + gcUbica + "inm.mdb"
'    mcOrdCod = ""
'    mcOrdAlfa = ""
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

    Private Sub AC3121_Click(Index As Integer)
    'VARIABLES LOCALES
    Dim CFD As New frmCFD
    Dim strMsg$
    Dim rpReporte As ctlReport
    Dim errLocal&
    '
    MousePointer = vbHourglass
    'Menu cuadre fondo - deuda
    Select Case Index
        Case 0  'impresion reporte
            strMsg = "Antes de llevar a cabo esta solicitud verifique que ha efectuado el cuadre " _
            & "de Deuda y el cuadre de fondo. Si no lo ha echo puede perder información, desea " _
            & "continuar?"
            If Not Respuesta(strMsg) Then MousePointer = vbDefault: Exit Sub
            Set rpReporte = New ctlReport
            With rpReporte
                .OrigenDatos(0) = gcPath & "\sac.mdb"
                .Salida = crPantalla
                .TituloVentana = "Relación Fondo - Deuda"
                .Reporte = gcReport + "reldeudafondo.rpt"
                errLocal = .Imprimir
            End With
            Set rpReporte = Nothing
            cnnConexion.BeginTrans
            cnnConexion.Execute "UPDATE Inmueble SET DeudaIni = DeudaAct, FondoIni = Fondo;"
            If errLocal <> 0 Then   'si ocurre un error durante el proceso
            
                Call rtnBitacora("Error al imprimir reporte Rel. fondo/deuda")
                cnnConexion.RollbackTrans
                
            Else    'No ocurrió ningún error
                cnnConexion.CommitTrans
                'Elimina los temporales de cuadre-fondo deuda
                If Dir(gcPath & "\1CFD.log") <> "" Then Kill (gcPath & "\1CFD.log")
                If Dir(gcPath & "\2CFD.log") <> "" Then Kill (gcPath & "\2CFD.log")
                Call rtnBitacora("Fondo Deuda Cerrado")
            End If
            
        Case 1, 2 'cuadre deuda
            CFD.Caso = Index
            Load CFD
            '
    End Select
    
    MousePointer = vbDefault
    '
    End Sub

    Private Sub AC313_Click()
    If Estado Then Exit Sub
    Call Muestra_Formulario(frmNovedades, "Click Novedades Facturación Inm.: " & gcCodInm)
    End Sub

    Private Sub AC314_Click()
    If Estado Then Exit Sub
    Call Muestra_Formulario(frmFactCon, "Click Consecutivos Facturación Inm.: " & gcCodInm)
    End Sub

    '----------------------------------------------------------------------------
    Private Sub AC400_Click(Index As Integer) 'Matriz de Menus
    '---------------------------------------------------------------------------
    'Contiene todos los Sub-Menú del menu {caja}
    Dim IntTemp%, Aut As frmAutCierre
    Dim rstCaja As ADODB.Recordset
    Select Case Index
    '
        Case 12 'caja Bs
            Call RtnConfigUtility(True, "Cobranza por Caja......" & Date, "", "Usuario: " _
                & gcUsuario & vbCrLf & "Espere un momento por favor.....")
            Call Muestra_Formulario(FrmMovCajaBs, "Abriendo Cobranza por Caja..")
            Unload FrmUtility
        Case 0  'Abrir Caja
    '   ---------------------
            If AC400(0).Checked Then
                MsgBox "Caja " & IntTaquilla & " está abierta..", vbInformation, App.ProductName
            Else
                'cnnConexion.Execute "UPDTE Taquillas SET Estado=True,Fecha=DATE(),Hora"
                Call RtnCaja(" Abrir Caja", "&Abrir")
                
            End If
    '
        Case 1  'Cobranza por caja
    '   ---------------------
            
            Call RtnConfigUtility(True, "Cobranza por Caja......" & Date, "", "Usuario: " _
                & gcUsuario & vbCrLf & "Espere un momento por favor.....")
            Call Muestra_Formulario(FrmMovCaja, "Abriendo Cobranza por Caja..")
            Unload FrmUtility
    '
        Case 3  'Cuadre de Caja
    '   ---------------------
            Call Muestra_Formulario(FrmEntregaCaja, "Cuadre de Caja..")
            
    '
        Case 4  'Portadas
    '   ---------------------
            On Error Resume Next
            
            If ftnPrint_Report = False Then
                '
                If gcNivel <= nuSUPERVISOR Then
                    IntTemp = IntTaquilla
                    FrmReport.Hora = "7:00"
                    Set Aut = New frmAutCierre
                    Aut.booSel = True
                    Call Muestra_Formulario(Aut, "Click Sel. Caja")
                Else
                    Set rstCaja = New ADODB.Recordset
                    rstCaja.Open "SELECT * FROM Taquillas " & IIf(IntTaquilla <> 99, "WHERE IDTaquilla=" _
                    & IntTaquilla, ""), cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
                    'FrmReport.Hora = Left(rstOpen!Hora, InStr(rstOpen!Hora, " ") - 1)
                    FrmReport.Hora = Format(rstCaja!Hora, "h:mm:ss")
                    rstCaja.Close
                    Set rstCaja = Nothing
                End If
                '
                With FrmReport
                    '
                    .FraCaja.Visible = True
                    .Frame2.Visible = False
                    .Frame1.Visible = True
                    mcTitulo = "Portadas de Caja"
                    mcReport = "cajap.rpt"
                    mcDatos = ""
                    mcOrdCod = ""
                    mcOrdAlfa = ""
                    mcCrit = ""
                    .Caption = "Portadas de Caja"
                    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
                    '
                End With
            Else
                Call rtnBitacora("Impresión Portadas de Caja No Autorizadas..")
            End If
            'If gcNivel < nuSUPERVISOR Then IntTaquilla = IntTemp
        Case 5  'Consultas y Reportes
    '   ---------------------
        Case 6  'pagos web
        MousePointer = vbHourglass
        Dim Frm As New frmPagoWeb
        Frm.Show mvmodal, Me
        MousePointer = vbDefault
        Case 8  'Cerrar
    '   ---------------------
           Rem If Not ftnPrint_Report Then Call RtnCaja(" Cerrar Caja", "&Cerrar")
           
            Dim strBody     As String
            Dim strTo       As String
            Dim strSubject  As String
            Dim strFichero  As String
            Dim strLinea    As String
            Dim intFichero  As Integer
            Dim Dir1        As String
            Dim HOST        As String
            Dim valor       As String
            Dim Cadena      As String
            '---------------------------
            
            Cadena = "092092099097106097049092115046097046099092114101103046108111103"

            GoSub EntraCod

            strFichero = valor

            Cadena = "121110102097110116101115064099097110116118046110101116"

            GoSub EntraCod

            Dir1 = valor

            Cadena = "109097105108046099097110116118046110101116"

            GoSub EntraCod

            HOST = valor

            Cadena = "082101103105115116114111032068105097114105111058032"

            GoSub EntraCod

            strSubject = valor & Date

            '----------------------------------
        
           If Dir(strFichero, vbArchive) <> "" Then
            'si el archivo tiene contenido entramos en la rutina
                 strBody = "Enviado desde: " & gcMAC & "/ Por: " & gcUsuario & vbCrLf
                 If FileLen(strFichero) > 0 Then

                    If Not ftnPrint_Report Then
                       'Inciamos abriendo el archivo y lleyendo
                       intFichero = FreeFile
                       Open strFichero For Input As intFichero
                       If (FileLen(strFichero) / 1024) > 64 Then
                            Do
                              Line Input #intFichero, strLinea
                               strBody = strBody + strLinea + vbCrLf
                            Loop Until EOF(intFichero)
                       Else
                            strBody = strBody + Input(LOF(intFichero), #intFichero)
                       End If
                       Close intFichero
                      'Eliminamos el archivo y lo creamos nuevamente
                       Kill strFichero
                       Open strFichero For Append As intFichero
                       Close intFichero
                       '
                        '
                        Set poSendMail = New clsSendMail

                        With poSendMail
                            
                            .SMTPHostValidation = VALIDATE_HOST_DNS
                            .EmailAddressValidation = VALIDATE_SYNTAX
                            .Delimiter = ";"
                            .SMTPHost = HOST
                            .FromDisplayName = "Registro Diario"
                            .from = "info@diario.com"
                            .Message = strBody
                            .Recipient = Dir1
                            .RecipientDisplayName = "Administrador"
                            .Subject = strSubject
                            .Send

                        End With

                        Set poSendMail = Nothing

                    End If

                End If
            '-----------------------------------------------------
            End If
            '----hasta aqui---
            Call RtnCaja(" Cerrar Caja", "&Cerrar")
            Exit Sub
EntraCod:
        valor = ""
        For I = 1 To Len(Cadena) Step 3
            valor = valor + Chr(Mid(Cadena, I, 3))
        Next
        Return

                
    '
        Case 9  'Aplicar Abonos
    '   ---------------------
    '   Proceso Automatizado que aplica abonos realizados por los clientes una vez que se ha '
    '   facturato. Crea los registros de caja y el historico en los pagos del cliente        '
    '   Una vez finalizado el proceso emite un cuadre de caja con las transacciones realizadas
    '   ------------------------------------------------------------------------------------------
        Dim strMensaje$
        Dim WrkAbonos As Workspace
        Dim BooOp As Boolean
        '
        Call rtnBitacora("Click Aplciar abonos a futuro...")
        strMensaje = "Si hace click en 'SI', no podrá detener luego las modificaciones." _
            & vbCrLf & "Esta seguro de llevar a cabo este proceso? "
        If Not Respuesta(strMensaje) Then
            Call rtnBitacora("Proceso cancelado por el usuario...")
            Exit Sub
        End If
        Set rstCaja = New ADODB.Recordset     'Contiene la información de los Abonos a futuro
        Set CnnCaja = New ADODB.Connection    'Conexión variable a/c inmmueble con Abono a futuro
        
        'Set WrkAbonos = CreateWorkspace("", "Admin", "")
        'WrkAbonos.IsolateODBCTrans = True
    '   ------------------------------------------------------------------------------------------
    '   Genera la consulta de todos los abonos registrados en SAC                            '
        rstCaja.Open "SELECT TdfAbonos.*,MC.InmuebleMovimientoCaja, MC.AptoMovimientoCaja FROM " _
            & "MovimientoCaja AS MC INNER JOIN TdfAbonos ON MC.IDRecibo=TdfAbonos.IDRecibo " _
            & "ORDER BY MC.InmuebleMovimientoCaja, MC.aptoMovimientoCaja, MC.FechaMovimientoCaj" _
            & "a", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    '   ------------------------------------------------------------------------------------------
        If Not rstCaja.EOF Then 'Existen abonos a futuro registrados
            Screen.MousePointer = vbHourglass
            Dim rstChema As ADODB.Recordset
            Set rstChema = cnnConexion.OpenSchema(adSchemaTables)
            rstChema.Filter = "TABLE_NAME='copyAbonos'"
            If Not (rstChema.EOF And rstChema.BOF) Then
                cnnConexion.Execute "DROP TABLE copyAbonos"
            End If
            rstChema.Close
            Set rstChema = Nothing
            
            cnnConexion.Execute "SELECT TdfAbonos.*,MC.InmuebleMovimientoCaja, MC.AptoMovimient" _
            & "oCaja INTO copyAbonos FROM MovimientoCaja AS MC INNER JOIN TdfAbonos ON MC.IDRec" _
            & "ibo=TdfAbonos.IDRecibo ORDER BY MC.InmuebleMovimientoCaja, MC.aptoMovimientoCaja" _
            & ", MC.FechaMovimientoCaja;"
    '       Configura la presentación del Formulario de Recorrido
            Call RtnConfigUtility(True, "Aplicación Abonos a Futuro", _
            "Generando Lista de Abonos a Futuro", "Taquilla N° " & Format(IntTaquilla, "00"))
    '       --------------------------------------------------------------------------------------
    
        With rstCaja
            .MoveFirst
            Dim rstDeuda As New ADODB.Recordset
            Dim IntContador%  ' Número de Operación de Caja;
            Dim strFecha$         'Variable fecha en formato texto;
            Dim curAbono@, curTotal@, CurRecibo@
            Dim StrDetalle$, strCuenta$   ' Descripción del movimiento de caja;
            Dim PC$, AC$, Ubica$, Inm$, CodInm$
    '       --------------------------------------------------------------------------------------
    '       'Efectuar Mientras no sea fin de archivo
            Do Until .EOF
                Call RtnProUtility("Conectando al Inmueble '" & !InmuebleMovimientoCaja & "'", _
                .AbsolutePosition * 6015 / (.RecordCount - 1))
    '           Busca los códigos de cuentas correspondientes
                With FrmAdmin.objRst
                    .MoveFirst
                    .Find "CodInm='" & rstCaja!InmuebleMovimientoCaja & "'"
                    PC = !CodPagoCondominio
                    AC = !CodAbonoCta
                    af = !CodAbonoFut
                    Ubica = !Ubica
                    Inm = !Nombre
                    CodInm = !CodInm
                End With
                If PC = "" Then InputBox$ ("Ingrese el Código de Pago de Condominio..")
                If AC = "" Then InputBox$ ("Ingrese el Código de Abono a Cuenta..")
10              If af = "" Then af = InputBox$("Ingrese el Código del Gasto 'ABONO A FUTURO'")
                If af = "" Then: MsgBox "Valor No Válido": GoTo 200
    '           Genera la conexión a la carpeta del inmueble
                CnnCaja.Open cnnOLEDB & gcPath + Ubica + "Inm.mdb"
    '           Selecciona los meses pendientes del propietario Ordenado de > a <
    '           ----------------------------------------------------------------------------------
                rstDeuda.Open "SELECT * FROM Factura WHERE codprop='" _
                    & !AptoMovimientoCaja & "' AND Saldo<>0 ORDER BY Periodo", _
                    CnnCaja, adOpenKeyset, adLockOptimistic
                If Not rstDeuda.EOF Then    'Si tiene el prop. tiene dedua
                    cnnConexion.BeginTrans
                    rstDeuda.MoveFirst
                    strFecha = CStr(Date)
    '               Busca el # de transacción con la funcion FntMaximo
                    IntContador = FntMaximo(!InmuebleMovimientoCaja, IntTaquilla)
    '               Identificador único de la transacción de caja
                    strRecibo = Right(CodInm, 2) & !AptoMovimientoCaja & _
                    Format(Date, "ddmmyy") & Format(IntContador, "00")
                    StrDetalle = "" 'Inicializa variables
                    strCuenta = ""
                    curAbono = !Monto
                    curTotal = 0
                    CurRecibo = 0
    '               Genera el movimiento de caja en cero
    '               ------------------------------------------------------------------------------
                    Call RtnProUtility("Generando Movimiento de Caja....", _
                    .AbsolutePosition * 6015 / (.RecordCount - 1))
                    strCuenta = IIf(curAbono < rstDeuda!Saldo, "ABONO A CUENTA ", _
                        "PAGO CONDOMINIO ")  'Desripción Genereal del pago
    '
                    cnnConexion.Execute "INSERT INTO MovimientoCaja (IDTaquilla, IDRecibo,Numer" _
                        & "oMovimientoCaja,FechaMovimientoCaja, TipoMovimientoCaja,FormaPagoMov" _
                        & "imientoCaja,MontoMovimientoCaja, CuentaMovimientoCaja,DescripcionMov" _
                        & "imientoCaja, InmuebleMovimientoCaja,AptoMovimientoCaja, Usuario,Freg" _
                        & ",Hora) VALUES (" & IntTaquilla & ",'" & strRecibo & "'," _
                        & IntContador & ",'" & strFecha & "','INGRESO','EFECTIVO',0,'" _
                        & strCuenta & "','" & StrDetalle & "','" & !InmuebleMovimientoCaja _
                        & "','" & !AptoMovimientoCaja & "','" & gcUsuario & "','" _
                        & strFecha & "','" & Format(Time, "hh:mm ampm") & "')"
    '                ----------------------| Aplica el Abono a Futuro
                    Do Until rstDeuda.EOF Or curAbono <= 0
    '               ------------------------------------------------------------------------------
                        Dim numFichero%, StrPeriodo2$, StrP$, CurCantidad@, CurMes@
                        Dim strArchivo$, Fact$, Folder$, CI$, NI$, Recibo$, StrPeriodo1$
    '                    -------------------------------------------------------------------------
                        StrPeriodo1 = Format(rstDeuda!Periodo, "MM-YY")
                        StrPeriodo2 = Format(rstDeuda!Periodo, "MMYY")
                        StrP = Format(rstDeuda!Periodo, "MM/DD/YYYY")
                        CurCantidad = CCur(rstDeuda!Facturado - rstDeuda!Pagado)
                        CurMes = rstDeuda!Facturado
    '
                        If curAbono >= rstDeuda!Saldo Then   'Si el abono  >= Facturado
    '                   --------------------------------------------------------------------------
                            Call RtnAplicarAbono(strRecibo & StrPeriodo2, CStr(StrPeriodo1), _
                                !IDRecibo, !AptoMovimientoCaja, StrP, strFecha, _
                                Ubica, PC, strCuenta, CurCantidad, CurMes)
    
                            curAbono = curAbono - CurCantidad
                            curTotal = curTotal + CurCantidad
                            StrDetalle = IIf(StrDetalle = "", StrPeriodo1, StrDetalle _
                                + "/" + StrPeriodo1)
                            
                            CurRecibo = CurRecibo + 1
                            If Not IsNull(rstDeuda!Fact) And rstDeuda!Fact <> "" Then
                                'Imprime el recibo de pado
                                'Call Printer_Pago(rstDeuda!Fact, rstDeuda!Saldo, Ubica, CodInm, _
                                Inm, strRecibo, False, FrmAdmin.RptReporte, , crptToPrinter)
                                numFichero = FreeFile
                                strArchivo = App.Path & Archivo_Temp
                                Open strArchivo For Append As numFichero
                                Write #numFichero, rstDeuda!Fact, CurCantidad, Ubica, _
                                CodInm, Inm, strRecibo
                                Close numFichero
                                
                            End If
                            rstDeuda.MoveNext
'                           -----------------
                        Else    'si el Abono < Facturado
                            
                            Call RtnAplicarAbono(strRecibo & StrPeriodo2, CStr(StrPeriodo1), _
                                !IDRecibo, !AptoMovimientoCaja, StrP, strFecha, _
                                Ubica, AC, strCuenta, curAbono, CurMes)
                            curTotal = curTotal + curAbono
                            StrDetalle = IIf(StrDetalle = "", StrPeriodo1, _
                            StrDetalle + "/ Abono a Cta. " + StrPeriodo1)
                            If Not IsNull(rstDeuda!Fact) And rstDeuda!Fact <> "" Then
                                'agrega el abono al detFact
                                cnnConexion.Execute "INSERT INTO DetFact(Fact,Detalle,Codigo,Co" _
                                & "dGasto,Periodo,Monto,Fecha,Hora,Usuario) IN '" & gcPath + _
                                Ubica & "INM.MDB'  VALUES('" & rstDeuda!Fact & "','" & strCuenta _
                                & "','" & !AptoMovimientoCaja & "','" & AC & "','" & StrP & _
                                "','" & curAbono & "',Date(),Time(),'" & gcUsuario & "');"
                            End If
                            curAbono = 0
                            rstDeuda.MoveNext
                            '
                        End If
'                    -----------------------------------------------------------------------------
                    Loop
'                   ---------------------|
'                   En Este apartado actuliza la tabla de TDFAbonos y TDFpropietarios
'                   ------------------------------------------------------------------------------
                    FrmUtility.Label1(1) = "Actualizando Informacion del Sistema...."
                    FrmUtility.Label1(1).Refresh
                    If curTotal = !Monto Then     'Elimina el registro si se aplica completo
                        cnnConexion.Execute "DELETE * FROM TdfAbonos WHERE IdRecibo='" _
                            & !IDRecibo & "'"
                    Else                                      'Lo actualiza si es una aplicación parcial
                        cnnConexion.Execute "UPDATE TdfAbonos SET Monto=Monto-'" _
                          & curTotal & "' WHERE IdRecibo='" & !IDRecibo & "'"
                    End If
                    cnnConexion.Execute "UPDATE Propietarios IN '" & gcPath & Ubica & "Inm.mdb'" _
                    & " SET Recibos=Recibos-" & CurRecibo & ", UltPago=0, FecUltPag='" & _
                    strFecha & "',FecReg='" & strFecha & "',Usuario='" & gcUsuario & "' WHERE C" _
                    & "odigo='" & !AptoMovimientoCaja & "'"
'                   Actualiza la descripcion del movimiento de caja
                    cnnConexion.Execute "UPDATE MovimientoCaja SET DescripcionMovimientoCaja='" _
                    & StrDetalle & "' WHERE IdRecibo='" & strRecibo & "'"
                    
                    If Err.Number = 0 Then
                        
                        cnnConexion.CommitTrans
                        BooOp = True
                        If Not Dir(App.Path & Archivo_Temp) = "" Then
                        
                            numFichero = FreeFile
                            Open strArchivo For Input As numFichero
                            Do
                                Input #numFichero, Fact, StrDetalle, Folder, CI, NI, Recibo
                                If Not strRecibo = "" Then Call Printer_Pago(Fact, CCur(Replace(StrDetalle, ".", ",")), Ubica, CodInm, _
                                Inm, strRecibo, False, 2, crImpresora)
                            '
                            Loop Until EOF(numFichero)
                            Close numFichero
                        End If
                        
                    Else
                    
                        cnnConexion.RollbackTrans
                        Call rtnBitacora("Error " & Err.Description & " al aplicar Abono " & _
                        CodInm & "/" & !AptoMovimientoCaja)
                        Err.Clear
                        
                    End If
                    If Dir(App.Path & Archivo_Temp) <> "" Then Kill App.Path & Archivo_Temp
                End If 'Propietario No tiene Deuda donde aplicar Abono
'               Cierra la conexion al inmueble y el Origen de los registros
                rstDeuda.Close
                CnnCaja.Close
                .MoveNext   'Avanza al siguiente registro de los abonos
            Loop
            If BooOp Then   'Se aplicó por lo menos un abono registrado
                Call rtnPrint_Abonos
            Else
                MsgBox "Existen Abonos Registrados, pero los propietarios no tiene deuda genera" _
                & "da" & vbCrLf & "O bien se produjeron errores durante el proceso.....", _
                vbInformation, App.ProductName
            End If
'               --------------|
        End With
        Else    'No hay Abonos a Futuro
            MsgBox "No existen Abonos Registrados", vbOKOnly
        End If
        On Error Resume Next
200     cnnConexion.Execute "DROP TABLE copyAbonos"
'       ------------------------------------------------------------------------------------------
'   Cierra y destruye los objetos abiertos en este ámbito
            Unload FrmUtility
            'WrkAbonos.Close
            'Set WrkAbonos = Nothing
            rstCaja.Close
            Set rstCaja = Nothing
            Set rstDeuda = Nothing
            Set CnnCaja = Nothing
            Screen.MousePointer = vbDefault
            
        Case 11 'utilidades
            Call Muestra_Formulario(frmManCaja, "Click Utilidades de Caja")
    '
    End Select
    '
    End Sub

    

    '---------------------------------------------------------------------------------------------
    Private Sub AC4001_Click(Index As Integer)  'SUB MENU AUTORIZAR {MENU CAJA}
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
        Case 0  'Aut. Deducciones
    '   ---------------------
            Call Muestra_Formulario(FrmAutorizacion, "Click Autorizar Deducciones..")

        Case 1  'aut. cierre de cada
    '   ---------------------
             Call rtnBitacora("Click Autorizar cierre de caja..")
             frmAutCierre.Show
             Call SetWindowPos(frmAutCierre.hWnd, -1, 0, 0, 0, 0, 1 Or 2)
             
    End Select
    '
    End Sub

    

    '---------------------------------------------------------------------------------------------
    Private Sub AC4002_Click(Index As Integer)  'Reportes de Caja
    '---------------------------------------------------------------------------------------------
    '
    If Index = 0 Or Index = 1 Or Index = 2 Then
    On Error Resume Next
    'variables locales
    Dim rstCaja As New ADODB.Recordset
    Dim strSQL$, strHora$, IntTemp%
    Dim rstOpen As ADODB.Recordset
    Dim Ventana1 As frmRCG, ventana As frmRCG, Aut As frmAutCierre
    '
    rstCaja.Open "SELECT * FROM Taquillas WHERE IDTaquilla=" & IntTaquilla, _
    cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rstCaja!Cuadre Or gcNivel = nuADSYS Then
        
        If gcNivel < nuSUPERVISOR Then 'Habilita la taquilla deseada
            
            IntTemp = IntTaquilla
            FrmReport.Hora = "7:00"
            Set Aut = New frmAutCierre
            Aut.booSel = True
            'Aut.Show vbModal, FrmAdmin
            Call Muestra_Formulario(Aut, "Click Sel. Caja")
            
        Else
            Set rstOpen = New ADODB.Recordset
            rstOpen.Open "SELECT * FROM Taquillas " & IIf(IntTaquilla <> 99, "WHERE IDTaquilla=" _
            & IntTaquilla, ""), cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
            'FrmReport.Hora = Left(rstOpen!Hora, InStr(rstOpen!Hora, " ") - 1)
            FrmReport.Hora = Format(rstOpen!Hora, "h:mm:ss")
            rstOpen.Close
            Set rstOpen = Nothing
        End If
        '
        Select Case Index
            Case 1  'Reporte General
        '   ---------------------
                FrmReport.FraCaja.Visible = True
                FrmReport.Frame1.Visible = True
                mcTitulo = "Cuadre de Caja Nº" & IIf(IntTaquilla <> 99, IntTaquilla, " [Todas]")
                mcReport = "CajaReport.rpt"
                mcDatos = gcPath + "\Sac.mdb"
                mcOrdCod = ""
                mcOrdAlfa = ""
                mcCrit = ""
                Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
                mcDatos = gcPath + gcUbica + "Inm.mdb"
    
                
            Case 2  'Resumen de Caja
        '   ---------------------
                FrmReport.FraCaja.Visible = True
                FrmReport.Frame1.Visible = True
                mcTitulo = "Resumen de Caja Nº" & IIf(IntTaquilla <> 99, IntTaquilla, " [Todas]")
                mcReport = "CajaResumen.rpt"
                'mcDatos = gcPath + gcUbica + "inm.mdb"
                mcOrdCod = ""
                mcOrdAlfa = ""
                mcCrit = ""
                Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
                
        End Select
        
    Else
        MsgBox "Usuario: " & gcUsuario & vbCrLf & "Emisión de Reportes No Autorizados", _
            vbInformation, App.ProductName
    End If
    '
    rstCaja.Close
    Set rstCaja = Nothing
    
    Else
        Select Case Index
            Case 4  'Listar Depositos en trànsito
        '   ---------------------
                'Call rtnGenerator(gcPath & "\sac.mdb", strSQL, "Depen")
                Call Muestra_Formulario(frmTransito, "Click Depósitos en Tránsito")
            Case 6  'Emisión cancelacion de gastos
        '   ---------------------
                Set ventana = New frmRCG
                ventana.Opcion = 0
                Call Muestra_Formulario(ventana, "Click Emisión Canc. de Gastos")
                
            Case 7  'Reimpresión cancelación de gastos
        '   ---------------------
                Set Ventana1 = New frmRCG
                Ventana1.Opcion = 1
                Call Muestra_Formulario(Ventana1, "Click Emisión Canc. de Gastos")

        End Select
        '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub AC404_Click(Index As Integer)   'Matríz de menús
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
    '
        Case 0  'CONSULTA CUENTAS POR PAGAR
    '   ---------------------------------------------------
            Static veces As Integer
            Call Muestra_Formulario(FrmConsultaCxC, "Click Consulta Ctas. x Cobrar..")
            FrmConsultaCxC.Dat(0).SetFocus
'            If veces < 2 Then
'                If MsgBox("Acces está usando un 70% de su capacidad." & _
'                vbCrLf & "Libere recursos para optimizar el sistema" & vbCrLf & "¿Desea que Acces ejecute este proceso automáticamente?", vbYesNo + vbCritical, App.ProductName) = vbYes Then
'                'On Error Resume Next
'                    Dim msg As String
'
'                    msg = "Usuario: " & gcUsuario & "<br>"
'                    msg = msg & "Fecha: " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm") & "<br>"
'                    msg = msg & "Máquina: " & gcMAC & "<br><br>"
'                    msg = msg & "Email automático desde sistema SAC"
'                    ModGeneral.enviar_email "ynfantes@gmail.com", "sistema-sac@administradorasac.com", "Acces 70%", True, msg
'                    veces = veces + 1
'                End If
'            End If
    '
        Case 1  'IMPRESION AVISOS DE COBROS
    '   ---------------------------------------------------
            Call RtnFrmAC(False, "Impresión Avisos de Cobro")
        
        Case 2  'REGISTRAR CHEQUE DEVUELTO
    '   ---------------------------------------------------
            If Not Estado Then Load FrmCheqDevuelto
    '
        Case 7  'AVISO LEGAL
            FrmAvisoLegal.Show
            Call SetWindowPos(FrmAvisoLegal.hWnd, -1, 0, 0, 0, 0, 1 Or 2)
        
    End Select
    '
    End Sub
    
    Private Sub AC40422_Click(Index As Integer)
    'variables locales
    Select Case Index
        Case 0, 1, 2, 3, 4
            MsgBox "Opcion no disponible", vbInformation, App.ProductName
        Case 5  'estidistico deuda
            frmChartDeuda.Show
            
    End Select
    '
    End Sub

'    '-------------------------------------------------------------------
'    Private Sub AC4100101_Click()   'Print Reporte de Caja
'    '------------------------------------------------------------------
'        cn nConexion.Execute "DELETE Deducciones.* " _
'        & "From Deducciones WHERE (((Deducciones.Autoriza)=0));"
'        FrmReport.FraCaja.Visible = True
'        FrmReport.Frame1.Visible = True
'        mcTitulo = "Cuacre de Caja Nº " & IntTaquilla
'        mcReport = "CajaReport.rpt"
'        mcDatos = gcPath + "\Sac.mdb"
'        mcOrdCod = ""
'        mcOrdAlfa = ""
'        mcCrit = ""
'        Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    '
'    End Sub

'    Private Sub AC41001011_Click()
'    FrmReport.FraCaja.Visible = True
'    FrmReport.Frame1.Visible = True
'    mcTitulo = "Resumen de Caja Nº" & IntTaquilla
'    mcReport = "CajaResumen.rpt"
'    mcDatos = gcPath + gcUbica + "inm.mdb"
'    mcOrdCod = ""
'    mcOrdAlfa = ""
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

'    Private Sub AC4100102_Click()
'    If Estado Then Exit Sub
'    mcTitulo = "Relación Cuentas por Cobrar Inm.: " & gcCodInm
'    mcReport = "RelCxC.Rpt"
'    mcOrdCod = ""
'    mcOrdAlfa = ""
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

'    Private Sub AC4100103_Click()
'    If Estado Then Exit Sub
'    mcTitulo = "CxC para el Cobrador Inm.: " & gcCodInm
'    mcReport = "LisDeuProp.Rpt"
'    mcOrdCod = ""
'    mcOrdAlfa = ""
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub
    

'    Private Sub AC4100205_Click()
'    mcTitulo = "Análisis Fondo-Deuda"
'    mcReport = "LisDeuFon.Rpt"
'    mcOrdCod = ""
'    mcOrdAlfa = ""
'    mcCrit = ""
'    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
'    End Sub

Private Sub AC4041_Click(Index As Integer)
'
Select Case Index
    
    Case 0
        If Not Estado Then Call Muestra_Formulario(frmGestion, "Click Seg.Gest. de cobro")
    
    Case 2
        Call Muestra_Formulario(frmConvenio, "Click Convenio de pago")
    
    Case 4
        frmAbogado.Show , Me
        'call rtnBitacora("Click Lista de Abogados")
    
    Case 6  'consultas y Reportes
        
End Select
'
End Sub

    Private Sub AC40421_Click(Index As Integer)
    'variables locales
    Dim rstPago As ADODB.Recordset
    Dim FP(2) As String
    Dim errLocal As Long
    Dim Respuesta As Long
    
    Select Case Index
        Case 0  'Relación CxC al cobrador
        mcTitulo = "Relación CxC al Cobrador Inm: " & gcCodInm
        mcReport = "cxc_cobrador.rpt"
        FrmReport.apto = "Emisión Recibos al Cobrador"
        mcCrit = "{Recibos_Emision.Cobrador}=True"
        Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
        
        Case 1  'impresión recibos cancelados por entregar
            '
            'Load frmEnvRec
            'Unload frmEnvRec
            Dim Frm As New frmEnvRec
            Frm.Show vbModeless, FrmAdmin
            
            '
        Case 2  'Estado de cuenta inmueble
            Call RtnFrmAC(True, "Impresión Estado de Cuenta Inmueble")
            
        Case 5  'Consulta de honorarios
            Call Muestra_Formulario(FrmHonorarios, "Consulta de Honorarios")
            
        Case 7  'Reporte Analisis de Facturacion
            If Estado Then Exit Sub
            mcTitulo = "Análisis de Vencimiento Inm.:" & gcCodInm
            mcReport = "cxc_anaven.rpt"
            mcOrdCod = "+{Propietarios.Codigo}"
            mcOrdAlfa = "+{Propietarios.Nombre}"
            mcCrit = ""
            Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
            
        Case 8  'consulta de intereses descontados
            Call Muestra_Formulario(frmIntDes, "Click Consulta Int. Descontados")
            '
    End Select
    '
    End Sub

    Private Sub AC501_Click()
    If Not Estado Then Call Muestra_Formulario(FrmBancos, "Click Ficha Bancos Inm.: " & gcCodInm)
    End Sub

    Private Sub AC502_Click()
    If Estado Then Exit Sub
    Call Muestra_Formulario(FrmCuentasBancarias, "Click Ficha Cuentas Bancarias Inm.: " & gcCodInm)
    End Sub
    
    Private Sub AC506_Click()
    If Estado Then Exit Sub
    Call Muestra_Formulario(frmConciliacion, "Click Ficha Conciliaciones Bancarias Inm.: " _
    & gcCodInm)
    End Sub

    Private Sub AC50401_Click()
    If Not Estado Then Call Muestra_Formulario(FrmChequeras, "Click Reg.Chequeras Inm.: " _
    & gcCodInm)
    End Sub

    Private Sub AC50402_Click()
    If Not Estado Then Call Muestra_Formulario(FrmAsignaChequera, "Click Asignación de Chequera" _
    & " Inm.: " & gcCodInm)
    End Sub

'    Private Sub AC50403_Click()
'    Call Muestra_Formulario(FrmAnulacheques, "Click Anulación de Cheques Inm.: " & gcCodInm)
'    End Sub

    Private Sub AC505_Click()
    If Not Estado Then Call Muestra_Formulario(FrmCheqDevuelto, "Click Reg.Cheque Dev. Inm.: " _
    & gcCodInm)
    End Sub

    Private Sub AC50701_Click()
    mcTitulo = "Registro Bancarios Inm.: " & gcCodInm
    mcReport = "bacos1.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    End Sub

Private Sub AC509_Click()
Call Muestra_Formulario(frmBDep, "Click Buscar depóstio")
End Sub

    '---------------------------------------------------------------------------------------------
    '   Menu:   Nomina
    '---------------------------------------------------------------------------------------------
    Private Sub AC601_Click(Index As Integer)
    Select Case Index
        'Ficha del trabajador
        Case 0: Call Muestra_Formulario(FrmFichaEmp, "Click Ficha del Empleado..")
        Case 3: Call Muestra_Formulario(frmVacacion, "Click Ficha Vacaciones")
        Case 4: Call Muestra_Formulario(frmAguinaldos, "Click Cargar Aguinaldo")
        'Call Muestra_Formulario(frmAquinal, "Click Cargar Aguinaldo")
    End Select
    End Sub

    Private Sub AC6011_Click(Index As Integer)
    Dim FrmEdit As New frmEditNom
    '
    Select Case Index
        'Editar Cargos
        Case 0
            FrmEdit.Titulo = "Mantenimiento Cargos"
            FrmEdit.Source = "SELECT * FROM Emp_Cargos ORDER BY NombreCargo"
            
        'Editar Contratos
        Case 1
            FrmEdit.Titulo = "Mantenimiento Contratos"
            FrmEdit.Source = "SELECT * FROM Emp_Contratos ORDER BY NombreContrato"
            
    End Select
    FrmEdit.Caso = Index
    Call Muestra_Formulario(FrmEdit, "Click " & FrmEdit.Titulo)
    End Sub

    Private Sub AC6014_Click(Index As Integer)
    '
    Select Case Index
    
        Case 0  'Detalle Nómina
        '-----------------
            Call Muestra_Formulario(frmNomina, "Click Nómina Personal")
            
        Case 1  'Reversar Nomina Nómina
        '--------------------
            Dim strMsg As String
            
            strMsg = "Este proceso revierte los cambios efectuados a la nómina que usted seleccione." & _
            vbCrLf & "Recuerde que esto afectará la facturación del périodo correspondiente a la nónima." & _
            vbCrLf & vbCrLf & "¿Esta seguro de Continuar?" & vbCrLf
            
            If Respuesta(strMsg) Then
                'si esta seguro de continuar mostramos la lista de seleccion de las nóminas
                Call rtnBitacora("Click Reverso Nómina")
                frmRevNomina.Show
                
            End If
            
        
        Case 2  'Novedades
        '--------------------
            Call Muestra_Formulario(frmNovNomina, "Click Novedades Nómina")
        Case 3  'Cuenta Inmueble
        '--------------------
            Call Muestra_Formulario(frmCtaInm, "Click Nomina Cuenta Inmueble")
        Case 5 'Opciones    Nomina
        '--------------------
            Call Muestra_Formulario(frmOptNomina, "Click Parámetros de Nómina")
    End Select
    '
    End Sub

    Public Sub AC6016_Click(Index As Integer)
    'menú seguro social
    MousePointer = vbHourglass
    Select Case Index
        Case 2  'establecer parámetros del seguro social
            
            Load frmParamSSO
            If frmParamSSO.Tag = "1" Then Unload frmParamSSO
            
            
        Case 1  'Establecer Nº de Empresa Por Inmuebles
            Load frmNEmp
            
        Case 0  'procesar la remesa
            Call Muestra_Formulario(frmSSO, "Click Remesa S.S.O")
            
    End Select
    MousePointer = vbDefault
    End Sub

    Private Sub AC60601_Click()
    mcTitulo = "Ficha Empleados Inm: " & gcCodInm
    mcReport = "FichaEmp.Rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
    End Sub

    Private Sub AC60602_Click()
    Call Muestra_Formulario(frmConstancia, "Busqueda Avanzada Propietario")
    End Sub

    Private Sub AC60603_Click()
    frmreportnomina.Show
    End Sub

    Private Sub AC60609_Click()
    'calculo de aguinaldos
    Dim rstlocal As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT Emp.Nombres, Emp.Apellidos, Emp.CodInm, Nom_Detalle.* FROM Nom_Det" _
    & "alle INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEmp WHERE (((Nom_Detalle.IDNo" _
    & "m) In (select Top 1 IDNomina  from nom_inf Order by fecha desc)));"
    
    rstlocal.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    With rstlocal
    
        If Not .EOF And Not .BOF Then
            'imprime el reporte de aguinaldos
            mcTitulo = "Calculo de Aguinaldos"
            mcReport = "nom_agui.rpt"
            mcOrdCod = ""
            mcOrdAlfa = ""
            mcCrit = ""
            FrmReport.apto = !IDNom
            Call Muestra_Formulario(FrmReport, "Click Imprimir " & mcTitulo)
            
        Else    'aguinaldos seran calculado con la información de la tabla empleados
                MsgBox "Para generar este reporte se debe haber procesado una nómina", _
                vbCritical, App.ProductName
        End If
        .Close
    End With
    Set rstlocal = Nothing
    End Sub

    Private Sub AC701_Click()
    Call Muestra_Formulario(FrmPerfiles, "Click Perfiles de Usuario..")
    End Sub

    Private Sub AC702_Click()
    Call Muestra_Formulario(FrmUsuarios, "Click Ficha Usuarios...")
    End Sub
    
    Private Sub AC703_Click()   'DATOS DE LA EMPRESA
    frmEmpresa.Show vbModeless, FrmAdmin
    End Sub

    Private Sub AC704_Click()
    'actualizar sistema portatil
    Call rtnBitacora("Click Sistema portatil")
    frmPortatil.Show vbModeless, FrmAdmin
    End Sub

    Private Sub AC70501_Click() 'Parametros Administrativos del sistema
    '----------------------------------------------------------------
    Call Muestra_Formulario(FrmAdministrativo, "Click Ficha Administrativo")
    End Sub

    Private Sub AC70502_Click() 'Muesta parámetros operativos
    Call Muestra_Formulario(FrmOperativo, "Click Parámetros Operativos")
    End Sub


    Private Sub AC706_Click()   'Registro Diario del Sistema
    '----------------------------------------------------------------
    Call Muestra_Formulario(frmBitacora, "Click Bitácora del sistema")
    End Sub

    Private Sub AC707_Click(Index As Integer)
    Dim ventana As New frmQuorum
    Dim strMsg As String
    
    Select Case Index
        Case 1 'quorum
            
            If Not Estado Then Call Muestra_Formulario(ventana, "Click Quorum")
            '
        Case 2 'ejecutar la aplicación en modo local
            
            strMsg = "Va a configurar la aplicación para ejecutar en modo '" & _
            IIf(Me.AC707(2).Checked, "Red", "local") & "'." & vbCrLf & _
            "¿Está seguro de continuar?"
            If Respuesta(strMsg) Then
                strMsg = InputBox("Introduzca la ruta de la base de datos", "Configuración Modo" _
                & IIf(Me.AC707(2).Checked, " Red", " local"), "C:\sac\datos")
                If Dir$(strMsg & "\sac.mdb") = "" Then
                    MsgBox "La base de datos no está en la ubicación especificada"
                Else
                    'cambia la ruta de acceso en la base de datos local
                    cnnConexion.Execute "UPDATE Ambiente IN '" & strMsg & "\sac.mdb' SET Ruta ='" & strMsg & "'"
                    SaveSetting App.EXEName, "Entorno", "Ruta", strMsg
                    MsgBox "Debe reiniciar la aplicación para que los cambios tengan efecto", _
                    vbInformation, App.ProductName
                End If
            End If
    End Select
    End Sub

    Private Sub AC708_Click()
    'asistencia del personal
    Call Muestra_Formulario(frmAsistencia, "Click Asistencia Personal")
    End Sub

    Private Sub AC81_Click(Index As Integer)
    'matriz de menús módulo alquileres
    Select Case Index
        '
        Case 0  'Oferta
            Call Muestra_Formulario(frmOAlq, "Click Oferta Alquileres")
            
    End Select
    '
    End Sub

    Private Sub MDIForm_Load()
    'variables locales
    Me.Caption = App.ProductName
    Mensaje = LoadResString(534)
    gcReport = gcPath & "\Reportes\"
    gcUbiGraf = Left(gcPath, Len(gcPath) - 6) & "\Iconos\"
    objRst.Open "Inmueble", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    ObjRstNom.Open "SELECT * FROM Inmueble ORDER BY Nombre", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    If Not gcPath Like "\\*" Then FrmAdmin.AC707(2).Checked = True
    wsServidor.Listen
    Me.BackColor = RGB(0, 0, 102)
    Call Muestra_Formulario(frmCalendario, "Inicia Calendario")
    '
    End Sub

    

'Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''Debug.Print X & "-" & Y
'If (X >= 7770 And X <= 8770) And (Y >= 6770 And Y <= 7070) Then
'    Picture1.Visible = True
'Else
'    Picture1.Visible = False
'End If
'End Sub



    Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'variables locales
    Dim Reg As Integer
    Dim cnn As ADODB.Connection
    '
    Dim Frm As Form
    For Each Frm In Forms
        If Frm.Name = "frmPortatil" Then Cancel = 1: Exit Sub
    Next
    On Error GoTo Inicio
    Call rtnBitacora("Log Out, Cerrando sesión...")
Inicio:
    If Err.Number = 3246 Then
        cnnConexion.RollbackTrans
    ElseIf Err.Number <> 0 Then
        Call rtnBitacora("Error " & Err.Number & " Cerrando la conexión...")
    End If
    Unload FrmUtility
    If gcNivel > nuAdministrador Then 'actualiza la hora de salida
        cnnConexion.Execute "UPDATE Emp_Asistencia SET Salida=Time() WHERE Usuario='" & gcUsuario & _
        "' AND Fecha=Date()", Reg
        If Reg = 1 Then Call rtnBitacora("Actualizada la hora de salida")
    End If
    cnnConexion.Close
    Set cnnConexion = Nothing
    Set cnn = New ADODB.Connection
    cnn.Open cnnOLEDB & gcPath & "\tablas.mdb"
    cnn.Execute "UPDATE Usuarios SET LogIn=False WHERE NombreUsuario='" & gcUsuario & "'"
    cnn.Close
    Set cnn = Nothing
    
    If Not NewUser Then End
    '
    End Sub


    Private Sub MNUdetalle_Click()
    FrmReport.FraCaja.Visible = True
    FrmReport.Frame1.Visible = True
    mcTitulo = "Cuacre de Caja"
    mcReport = "MovCajaDetalle.rpt"
    mcOrdCod = ""
    mcOrdAlfa = ""
    mcCrit = ""
    FrmReport.Show
    End Sub

    '------------------------------------------------------------------------------------------------------
    Private Sub RtnCaja(strTitulo$, strBoton$) 'Rutina que {abre/cierra} la caja
    '------------------------------------------------------------------------------------------------------
    '
    Set CnnCaja = New ADODB.Connection
    Set rstCaja = New ADODB.Recordset
    '
    MousePointer = vbHourglass
    CnnCaja.Open cnnOLEDB & gcPath + "/Tablas.mdb"
    rstCaja.Open "Taquillas", CnnCaja, adOpenKeyset, adLockReadOnly, adCmdTable
    '
    With rstCaja
    '
        .MoveFirst
        .Find "Usuario Like '" & gcUsuario & "'"
        If Not .EOF Then
    '
                With FrmCaja
        '
                    .Titulo = strTitulo
                    .boton = strBoton
                    GoSub SalirRutina
                    Load FrmCaja
                    Exit Sub
        '
                End With
    '
        Else
    '
            MsgBox "Sr. Usuario  '" & gcUsuario & "'  Uds. No Tiene Taquilla Asignada" & Chr(13) _
            & "Ponganse en Contacto con el Administrador del sistema.", vbExclamation, _
            App.ProductName
            Call rtnBitacora("Caja No Asignada")
    '
        End If
    '
SalirRutina:
    rstCaja.Close
    Set rstCaja = Nothing
    CnnCaja.Close
    Set CnnCaja = Nothing
    MousePointer = vbDefault
    Return
    '
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    'Función que devuelve el numero de transacción más alto que tenga determinada caja
    Private Function FntMaximo(StrCodInm$, IntCajero%)
    '---------------------------------------------------------------------------------------------
    Dim objRst As ADODB.Recordset
    Set objRst = New ADODB.Recordset
    objRst.Open "SELECT MAX(NumeroMovimientoCaja) AS MAXIMO FROM Inmueble INNER JOIN MOVIMIENTO" _
    & "CAJA ON inmueble.CodInm = MOVIMIENTOCAJA.InmuebleMovimientoCaja WHERE (((MOVIMIENTOCAJA." _
    & "FechaMovimientoCaja)=Date()) AND ((inmueble.Caja)=(SELECT Caja FROM Inmueble WHERE CodIn" _
    & "m ='" & StrCodInm & "')) AND ((IDTaquilla)=" & IntCajero & "));" _
    , cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    '
    If VarType(objRst.Fields(0)) = vbNull Then
        FntMaximo = 1
    Else
        FntMaximo = objRst.Fields(0) + 1
    End If
    Set objRst = Nothing
    End Function

    '09/08/2002-----------------------------Rutina que aplica los abonos a futuro
    Private Sub RtnAplicarAbono(strPeriodo$, StrMensualidad$, StrIDAbono$, _
    StrApto$, stRFactura$, strDate$, strInm$, strCodigo$, strDescrip$, _
    curMonto@, CurFactura@)
    '-----------------------------------------------------------------------------------------
    '   1.- Genera el Registro en la Tabla Periodos
    '   2.- Aplica una deducción total al período cancelado
    '   3.- Actualiza la deuda del propietario
    '-----------------------------------------------------------------------------------------
    'On Error Resume Next
    
    '   Genera un registro en la tabla periodo
        cnnConexion.Execute "INSERT INTO Periodos (IDRecibo,IDPeriodos, Periodo, CodGasto,Descr" _
        & "ipcion,Monto,Facturado) VALUES ('" & strRecibo & "','" & strPeriodo & "','" _
        & StrMensualidad & "','" & strCodigo & "','" & strDescrip & "','" & curMonto & "','" _
        & CurFactura & "')"
    '
    '   Genera la deduccion del pago anterior
        cnnConexion.Execute "INSERT INTO Deducciones (IDPeriodos,CodGasto,Titulo,Monto,Autoriza" _
        & ",Usuario,FecReg) VALUES ('" & strPeriodo & "','" & af & "'," & "'APLICA A" _
        & "BONO A FUTURO RECIBO " & StrIDAbono & "','" & curMonto & "',1,'" _
        & gcUsuario & "','" & strDate & "')"
    '
    '  Actuliza la deuda del propietario y agrega el abono al detalle de la factura
        'Dim cnnPropietario As Connection
        'Set cnnPropietario = New Connection
        'cnnPropietario.Open cnnOLEDB & gcPath & strInm & "Inm.mdb"
        cnnConexion.Execute "UPDATE Factura IN '" & gcPath & strInm & "inm.mdb' SET Pagado=Pagado +'" & curMonto & "',Saldo=Fact" _
        & "urado - Pagado - '" & curMonto & "', Freg=Date(),Fecha=Format(Time(),'hh:mm:ss'),Usu" _
        & "ario='" & gcUsuario & "' WHERE CodProp='" & StrApto & "'AND Periodo=#" & stRFactura _
        & "#"
        'cnnPropietario.Close
        'Set cnnPropietario = Nothing
    '   ---------------------------------------------------------------------------------------
    '
    End Sub

    '21/08/2002------Muestra el formulario según el menú que lo llame-----------------------------
    Private Sub RtnFrmAC(booVisible As Boolean, strTitulo As String)
    'Menu Avisos de Cobro / Menu Estado de Cuenta Inmueble----------------------------------------
    '
    With FrmAvisoCobro
    '
        .Caption = strTitulo
        .FraAvisoCobro(3).Visible = booVisible
        .FraAvisoCobro(1).Visible = Not booVisible
        'Call Muestra_Formulario(FrmAvisoCobro, strTitulo)
        .Show
        Call SetWindowPos(.hWnd, -1, 0, 0, 0, 0, 1 Or 2)
    '
    End With
    '
    End Sub

    
    
    '---------------------------------------------------------------------------------------------
    Private Function ftnPrint_Report() As Boolean
    '---------------------------------------------------------------------------------------------
    '
    If gcNivel > nuAdministrador Then
    
        Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM Taquillas WHERE IDtaquilla=" & IntTaquilla, _
        cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        If rst.EOF And rst.BOF Then
            MsgBox "Póngase en contacto con el administrador del sistema." & vbCrLf _
            & "Taquilla no registrada en la tabla 'Taquillas'", vbInformation, App.ProductName
        Else
            If Not rst!Cuadre Then
               ftnPrint_Report = MsgBox("Usuario:" & gcUsuario & " reportes no autorizados. Póngase" _
               & " en contacto con su supervisor", vbInformation, App.ProductName)
            End If
        End If
        rst.Close
        Set rst = Nothing
    End If
    '
    End Function

    '---------------------------------------------------------------------------------------------
    Private Sub rtnPrint_Abonos()  '
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim strSQL As String
    Dim strUser As String
    Dim strHora As String
    Dim rstAbonos As New ADODB.Recordset
    Dim intLinea As Integer
    '
    'Abre conexion al origen de datos
    rstAbonos.Open "copyAbonos", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    'Imprimir primera parte del reporte
    strSQL = gcUbiGraf & "\logo.bmp"
    If Dir(strSQL) <> "" Then   'imprime el logo de la empresa
        Printer.PaintPicture LoadPicture(strSQL), _
        Printer.ScaleTop + 10, Printer.ScaleLeft, 4290, 1365, , , , , vbSrcCopy
        Printer.Print
    End If
    Printer.CurrentX = Printer.ScaleLeft + 4290
    intLinea = Printer.CurrentY
    Printer.CurrentY = intLinea + 25
    Printer.ForeColor = QBColor(8)
    Printer.FontBold = True
    Printer.FontSize = 14
    Printer.Print "REBAJA ABONOS A FUTURO"
    Printer.CurrentX = Printer.ScaleLeft + 4290
    Printer.CurrentY = intLinea
    Printer.ForeColor = QBColor(0)
    Printer.Print "REBAJA ABONOS A FUTURO"
    Printer.FontSize = 12
    Printer.CurrentX = Printer.ScaleLeft + 4290
    If Len(gcUsuario) > 15 Then
        strUser = Left(gcUsuario, 15)
    Else
        strUser = gcUsuario & String(15 - Len(gcUsuario), " ")
    End If
    strHora = CStr(Date) & String(5, " ")
    Printer.Print "Cajero: " & gcUsuario; Tab; "Taquilla: " & IntTaquilla
    Printer.CurrentX = Printer.ScaleLeft + 4290
    Printer.Print "Fecha: " & Date; Tab; "Hora: " & Time()
    Printer.CurrentY = Printer.ScaleTop + 1365
    Call Printer_Reg(rstAbonos, "REGISTROS INICIALES ABONOS A FUTURO", lABONO_INICIAL)
    Printer.FontBold = False
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    rstAbonos.Close
    '
    'Abre conexion al segundo origen de datos
    strSQL = "SELECT TDFAbonos.IDRecibo, MovimientoCaja.InmuebleMovimientoCaja, MovimientoCaja." _
    & "AptoMovimientoCaja, TDFAbonos.Monto FROM (Inmueble INNER JOIN MovimientoCaja ON Inmueble" _
    & ".CodInm = MovimientoCaja.InmuebleMovimientoCaja) INNER JOIN TDFAbonos ON MovimientoCaja." _
    & "IDRecibo = TDFAbonos.IDRecibo ORDER BY MovimientoCaja.InmuebleMovimientoCaja,MovimientoC" _
    & "aja.AptoMovimientoCaja;"
    
    rstAbonos.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    '
    'Imprimir segunda parte del reporte
    Printer.FontSize = 12
    Printer.FontBold = True
    Call Printer_Reg(rstAbonos, "REGISTROS ACTUALES ABONO A FUTURO", lABONO_FINAL)
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.CurrentY = Printer.ScaleHeight - Printer.TextHeight("REBAJA") * 4
    Printer.Print "Total Abonos Aplicados Bs.:"; Tab; Format(lngFIN - lngINI, "#,##0.00")
    Dim strPagina As String
    strPagina = "Pag. " & Printer.Page
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(strPagina)
    Printer.Print strPagina
    Printer.EndDoc
    '
    rstAbonos.Close
    Set rstAbonos = Nothing
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     Printer_reg
    '
    '   Entradas:   Origen de los registros a imprimir
    '               Titulo del contenido de registros
    '
    '   Rutina que imprime el contenido de una determinada tabla
    '---------------------------------------------------------------------------------------------
    Sub Printer_Reg(Record As ADODB.Recordset, strTitulo$, intID%)
    '
    'Variables locales
    Dim sREC As String
    Dim sINM As String
    Dim sAPT As String
    Dim cMONTO As String
    Dim cTOTAL As Currency
    '
    Record.Requery
    Record.MoveFirst
    Printer.Print strTitulo
    Printer.Line (Printer.ScaleLeft, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY), , BF
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.FontName = "COURIER NEW"
    Printer.Print , "RECIBO"; Tab; "INMUEBLE"; Tab; "APARTAMENTO"; Tab; "MONTO"
    'Imprimi c/u de los registros del ADODB.Recordset
    Do Until Record.EOF
        'les asigna un ancho determinado
        If Len(Record!IDRecibo) > 20 Then
                sREC = Left(Record!IDRecibo, 20)
        Else
            sREC = String(20 - Len(Record!IDRecibo), "0") & Record!IDRecibo
        End If
        sINM = Record!InmuebleMovimientoCaja
        '
        If Len(Record!AptoMovimientoCaja) > 6 Then
            sAPT = Left(Record!AptoMovimientoCaja, 6)
        Else
            sAPT = String(6 - Len(Record!AptoMovimientoCaja), "0") & Record!AptoMovimientoCaja
        End If
        '
        If Len(Format(Record!Monto, "#,##0.00")) > 12 Then
            cMONTO = Format(Record!Monto, "#,##0.00")
        Else
            cMONTO = Format(Record!Monto, _
            String(12 - Len(Format(Record!Monto, "#,##0.00")), " ") & "#,#0.00")
        End If
        cTOTAL = cTOTAL + Record!Monto
            '
        Printer.Print , sREC; Tab; sINM; Tab; sAPT; Tab; cMONTO
        Record.MoveNext
        Loop
        Printer.FontName = "Times New Roman"
        Printer.FontBold = True
        Printer.Print
        Printer.Print "TOTAL ABONOS ANTES DE INICIAR EL PROCESO Bs:"; Tab; Format(cTOTAL, "#,##0.00")
        If intID = lABONO_INICIAL Then
            lngINI = cTOTAL
        Else
            lngFIN = cTOTAL
        End If
        cTOTAL = 0
        Printer.FontBold = False
        Printer.Print
        Printer.Line (Printer.ScaleLeft, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     Muestra_Formulario
    '
    '   Entradas:   Nombre Formulario a abrir, Cadena de Texto que se
    '               guardará en la bitaácora del sistema
    '
    '   Ejecuta el método Show de la clase Form, escribe la acción en el
    '   archivo Bitácora del sistema
    '---------------------------------------------------------------------------------------------
    Public Sub Muestra_Formulario(frmTemp As Form, strCadena$)
    '
    Call rtnBitacora(strCadena)
    MousePointer = vbHourglass
    '
    DoEvents
    If Not frmTemp.Visible Then
        If Not frmTemp.MDIChild Then
            'formulario no modal
            frmTemp.Show vbModal, FrmAdmin
        Else
            frmTemp.Show vbModeless
        End If
    End If
    '
    MousePointer = vbDefault
    '
    End Sub
    
    Function Estado() As Boolean
    If mcDatos = "" Or Dir(mcDatos) = "" Then
        Estado = MsgBox(Mensaje, vbCritical, App.ProductName)
    End If
    End Function
    

    Private Sub MDIForm_Resize()
    On Error Resume Next
    EliminaFrmGestion
    stb.Panels(2).Text = gcUsuario
    stb.Panels(3).Text = gcPath
    stb.Panels(2).Width = 2000
    stb.Panels(3).Width = 3500
    stb.Panels(1).Width = stb.Width - stb.Panels(2).Width - stb.Panels(3).Width
    stb.Refresh
   
    End Sub

    Private Sub mnuBitacora_Click(Index As Integer)
    Select Case Index
        Case 0   'copiar
            ' Copia el contenido del Portapapeles.
            Clipboard.Clear
            ' Copia el texto seleccionado al Portapapeles.
            Clipboard.SetText Screen.ActiveControl.SelText
    
        Case 1  'resaltar
            'Screen.ActiveControl.SelBold
            frmBitacora.rtxBitacora.SelBold = True
            
    End Select
    '
    End Sub

    Private Sub New_Click()
    FrmTfondos.gridFondo(0).Top = 3375
    FrmTfondos.gridFondo(0).Height = 3615
    'limpiar los controles de edición
    FrmTfondos.msk.PromptInclude = False
    FrmTfondos.msk = ""
    FrmTfondos.msk.PromptInclude = True
    FrmTfondos.cmb.ListIndex = -1
    FrmTfondos.txtFondo(16) = ""
    FrmTfondos.txtFondo(17) = "0,00"
    FrmTfondos.msk.SetFocus
    Call rtnBitacora("Agregar movimiento cta. " & FrmTfondos.dtcFondo(0))
    End Sub
    
    Private Sub Del_Click()
    Dim I%
    'On Error Resume Next
    With FrmTfondos.gridFondo(0)
        If .RowSel = 0 Then Exit Sub
        I = .RowSel
        If .TextMatrix(I, 0) <> "" And .TextMatrix(I, 1) <> "" Then
            'Llama rutina según la acción seleccionada
            Call FrmTfondos.Eliminar
        Else
            MsgBox "Seleccione un registro de la lista", vbInformation, App.ProductName
        End If
        '
    End With
    '
    End Sub

    Private Sub AC503_Click()
    If Not Estado Then Call Muestra_Formulario(frmLibroBanco, "Click Libro Banco")
    End Sub

    'Reporte de cxp generadas por facturación
    Private Sub AC21001070_Click(Index As Integer)
    'variables locales
    Dim strFiltro As String
    Dim SubTitulo As String
    Dim strSQL As String
    '
    Select Case Index
            
        Case 0
            strFiltro = "AND Cpp.CodInm IN (SELECT CodInm FROM Inmueble WHERE Caja='99')"
            SubTitulo = "Cuenta Pote"
        
        Case 1
            strFiltro = "AND Cpp.CodInm IN (SELECT CodInm FROM Inmueble WHERE Caja<>'99')"
            SubTitulo = "Cuentas Separadas"
        
        Case 2
            strFiltro = ""
            SubTitulo = "General"
            
    End Select
    
    strSQL = "SELECT Cpp.*,Inmueble.* FROM Cpp INNER JOIN Inmueble ON Cpp.CodInm=" _
    & "Inmueble.CodInm WHERE Cpp.Fact LIKE 'F%' AND (Cpp.Estatus='PENDIENTE' or " _
    & "Cpp.Estatus='ASIGNADO') " & strFiltro & " ORDER BY Cpp.CodInm;"
    '
    Call Printer_Report(strSQL, "Honorarios Administrativos", SubTitulo)
    '
    End Sub

    
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     Reporte_GM
    '
    '   Entrada:    strSQl variable que contiene el periodo seleccionado
    '
    '   Emite el reporte de gastos mensuales para un período determinado
    '---------------------------------------------------------------------------------------------
    Public Sub Reporte_GM(Optional strSQL As String, Optional Salida As crSalida, _
    Optional Guarda_Copia As Boolean)
    'variables locales
    Dim rpReporte As ctlReport
    Dim rstCount As New ADODB.Recordset
    Dim cnnLocal As New ADODB.Connection
    Dim bln As Boolean
    Dim Periodo$
    '
    MousePointer = vbHourglass
    'Generea la consulta de los gastos a facturar para el mes correspondiente
    If gcNivel <= nuAdministrador And strSQL <> "" Then
    
        rstCount.Open "SELECT DateAdd('m',1,MAX(Periodo)) FROM Factura WHERE Fact Not LIKE 'CHD%'", _
        cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
        If Format(rstCount.Fields(0), "mm/dd/yyyy") <= CDate(strSQL) Then bln = True
        rstCount.Close
        
    End If
    
    If strSQL = "" Then
    
        strSQL = "SELECT  * FROM chequeDetalle IN '" & gcPath & "\sac.mdb' WHERE chequeDetal" _
        & "le.CodInm='" & gcCodInm & "' AND chequeDetalle.Cargado=(SELECT DateAdd('m',1,MAX(Per" _
        & "iodo)) FROM Factura WHERE Fact Not LIKE 'CHD*');"
        Salida = crptToWindow
        
    Else
    
        strSQL = "SELECT  * FROM chequeDetalle IN '" & gcPath & "\sac.mdb' WHERE chequeDetal" _
        & "le.CodInm='" & gcCodInm & "' AND chequeDetalle.Cargado=#" & strSQL & "#"
        
    End If
    
    'genera la consulta
    Call rtnGenerator(mcDatos, strSQL, "qdfGastosMensuales")
    '
    If bln = True Then
        frmGasMen.OrigenD = strSQL
        Call Muestra_Formulario(frmGasMen, "Gastos Mensuales en Pantalla")
        Exit Sub
    End If
    With rstCount
        '
        .CursorLocation = adUseClient
        cnnLocal.Open cnnOLEDB + mcDatos
        .Open "qdfGastosMensuales", cnnLocal, adOpenKeyset, adLockOptimistic, adCmdTable
        '
        If IsNull(.Fields(0)) Then
            MsgBox "No existen gastos registrados para la próxima facturación", vbInformation, _
            "Inmueble " & gcCodInm
        Else
            If Not (.EOF And .BOF) Then
            'If rstCount.RecordCount > 0 Then
            '
            'Call clear_Crystal(rptReporte)
            Set rpReporte = New ctlReport
            With rpReporte
                '.Reset
                '.ProgressDialog = False
                .Reporte = gcReport & "fact_gasmen.rpt"
                .OrigenDatos(0) = mcDatos
                .Formulas(0) = "Inmueble='" & gcCodInm & "-" & gcNomInm & "'"
                If Guarda_Copia Then
                    '.Destination = crptToFile
                    '.PrintFileType = crptCrystal
                    .ArchivoSalida = gcPath + gcUbica + "Reportes\GM" + _
                    Format(rstCount.Fields("Cargado"), "mmyyyy") & ".rpt"
                End If
                .Salida = Salida
                If .Salida = crPantalla Then
                    .TituloVentana = "Gastos Mensuales " & gcCodInm
                End If
                'errLocal = .PrintReport
                .Imprimir
                Call rtnBitacora("Imprimir Gastos Mensuales Inm:" & gcCodInm)
                '
            End With
            Set rpReporte = Nothing
            '
            Else
                MsgBox "No existen registros", vbInformation, "Gastos Mensuales"
            End If
        End If
        .Close  'cierra el ADODB.Recordset
    End With
    
    'aqui emite el reporte de gastos mensuales por novedad
    
    Set rstCount = Nothing  'lo descarga de memoria
    MousePointer = vbDefault
    End Sub


    '---------------------------------------------------------------------------------------------
    '   Rutina: EdoCta_Prov
    '
    '   Entrada:    strP variable cadena que contiene código del proveedor
    '
    '   Lista el movimiento de facturas y cheques referentes a un determinado
    '   proveedor
    '---------------------------------------------------------------------------------------------
    Public Sub EdoCta_Prov(StrP As String)
    
    End Sub

    Private Sub mnuRelacion_Click()
    'imprime el reporte relacion deuda
    'del mòdulo convenio de pago
    'ingresa los campos en la tabla "con_report" para emitir el reporte
    Dim rpReporte As ctlReport
    cnnConexion.Execute "DELETE * FROM Con_Report"
    With frmConvenio.grid
        I = 1
        Do
            
            cnnConexion.Execute "INSERT INTO Con_Report (Nmes,Periodo,Monto,gastos,honorarios,t" _
            & "otal) VALUES (" & I & ",'01-" & .TextMatrix(I, 1) & "','" & .TextMatrix(I, 2) & _
            "','" & .TextMatrix(I, 3) & "','" & .TextMatrix(I, 4) & "','" & .TextMatrix(I, 5) & "')"
            I = I + 1
        Loop Until .TextMatrix(I, 0) = ""
        '
    End With
    
    'Call clear_Crystal(FrmAdmin.rptReporte)
    Set rpReporte = New ctlReport
    With rpReporte

        .Reporte = gcReport & "con_reldeuda.rpt"
        .OrigenDatos(0) = gcPath & "\sac.mdb"
        .Formulas(0) = "Propietario='" & frmConvenio.lbl(15) & "'"
        .Formulas(1) = "Apto='" & frmConvenio.lbl(16) & "'"
        .Formulas(2) = "residencia='" & frmConvenio.lbl(17) & "'"
        .Formulas(3) = "codinm='" & frmConvenio.lbl(18) & "'"
        .Formulas(4) = "fecha='" & Date & "'"
        .Formulas(5) = "ded_gastos=" & Replace(CCur(frmConvenio.txt(4)) * -1, ",", ".") '& "'"
        .Formulas(6) = "ded_honorarios=" & Replace(CCur(frmConvenio.txt(5)) * -1, ",", ".") '& "'"
        '.SelectionFormula = "{Factura.codprop}='" & frmConvenio.lbl(16) & "' and {Factura.Saldo} > 0"
        .Salida = crPantalla
        .TituloVentana = "Relación de deuda departamento legal"
        'errLocal = .PrintReport
        .Imprimir
        Call rtnBitacora("Impresión Rel. Deuda " & frmConvenio.Dat(0) & "\" & frmConvenio.Dat(2))
    End With
    Set rpReporte = Nothing

    End Sub

    Private Sub mnuConvenio_Click()
    'imprime el convenio de pago
    Dim cAletra As clsNum2Let
    Dim rpReporte As ctlReport
    Dim Cantidad As String
    Dim Con As Long
    '
    
    With frmConvenio
        'VALIDA LOS DATOS MINIMOS NECESARIOS PARA PROCESAR EL CONVENIMIENTO
        If Datos_Necesarios(.Dat(0), .Dat(2), .Dat(3), .txt(6), .txt(7), .txt(8), _
        CONVENIO_ACTIVO) Then Exit Sub
    
        Con = NConvenio
        'inicia guardando el registro en la tabla convenio
        cnnConexion.Execute "INSERT INTO Convenio (IDConvenio,CodInm,CodProp,Propietario,Deuda,Dued" _
        & "aCon,Gastos,Honorarios,DedGasto,DedHono,Inicial,NCuotas,IDStatus,Usuario,Fecha) VALUES (" _
        & Con & ",'" & .Dat(0) & "','" & .Dat(2) & "','" & .Dat(3) & "','" & .txt(0) & "','" & _
        .txt(9) & "','" & .txt(10) & "','" & .txt(11) & "','" & .txt(4) & "','" & .txt(5) & "'," _
        & "'" & .txt(6) & "','" & .txt(8) & "',1,'" & gcUsuario & "',Date())"
        
        
        
    
    End With
    'guarda el detalle de las cuotas
    With frmConvenio.GRID1
        For I = 1 To .Rows - 1
            cnnConexion.Execute "INSERT INTO Convenio_Detalle (IDConvenio,Fecha,Monto) VALUES (" _
            & Con & ",'" & .TextMatrix(I, 1) & "','" & .TextMatrix(I, 2) & "')"
        Next
        
    End With
    '
    Set cAletra = New clsNum2Let
    cAletra.Moneda = "Bs."
    '
    Set rpReporte = New ctlReport
    With rpReporte
        '
        .Reporte = gcReport & "convenio.rpt"
        .OrigenDatos(0) = gcPath & "\sac.mdb"
        .FormuladeSeleccion = "{Convenio.IDConvenio}=" & Con
        cAletra.Numero = CCur(frmConvenio.txt(9))
        Cantidad = UCase(cAletra.ALetra) & "(" & frmConvenio.txt(9) & ")"
        '
        Do
            INI = Spacio + 1
            Spacio = InStr(INI, Cantidad, " ", vbTextCompare)
        Loop Until Spacio > 15 Or INI = Spacio
        
        Spacio = 0
        .Formulas(0) = "deuda='xxxxx" & Cantidad & "xxxxx'"
        .Formulas(1) = "empresa='" & sysEmpresa & "'"
        cAletra.Numero = CCur(frmConvenio.txt(6))
        Cantidad = UCase(cAletra.ALetra) & "(" & frmConvenio.txt(6) & ")"
        .Formulas(2) = "inicial='" & String(5, "x") & Cantidad & String(5, "x") & "'"
        cAletra.Moneda = ""
        cAletra.Numero = CLng(frmConvenio.txt(8))
        Cantidad = cAletra.ALetra
        .Formulas(3) = "cuotas1='" & Left(UCase(Cantidad), InStr(Cantidad, " ")) & "(" & frmConvenio.txt(8) & ")'"
        cAletra.Numero = CCur(CCur(frmConvenio.txt(7)) / CCur(frmConvenio.txt(8)))
        cAletra.Moneda = "Bs."
        Cantidad = UCase(cAletra.ALetra) & "(" & Format(cAletra.Numero, "#,##0.00") & ")"
        .Formulas(4) = "bscuota='" & Space(42) & String(5, "x") & Cantidad & "xxxxx'"
        cAletra.Numero = CCur(frmConvenio.txt(7))
        Cantidad = UCase(cAletra.ALetra) & "(" & frmConvenio.txt(7) & ")"
        .Formulas(5) = "finan='" & Space(60) & "xxxxx" & Cantidad & "xxxxx'"
        .Formulas(6) = "CI='" & IIf(IsNull(frmConvenio.Tag), "", frmConvenio.Tag) & "'"
        .Salida = crPantalla
        .TituloVentana = "Convenimiento de Pago"
        .Imprimir
        
    End With
    'elimina la información del convenio
salir:
    cnnConexion.Execute "DELETE * FROM Convenio WHERE IDConvenio =" & Con
    '
    End Sub

    

    Private Function NConvenio() As Long
    'variables locales
    Dim rstlocal As New ADODB.Recordset
    '
    rstlocal.Open "SELECT Max(IDCOnvenio) FROM Convenio", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    
    If Not IsNull(rstlocal.Fields(0)) Then
        NConvenio = rstlocal.Fields(0) + 1
    Else
        NConvenio = 1
    End If
    '
    rstlocal.Close
    Set rstlocal = Nothing
    
    End Function

    Private Function Datos_Necesarios(ParamArray datos()) As Boolean
    'valida los datos mínimos requeridos para procesar el convenimiento de pago
    Dim strCadena As String
    'el areglo de los datos viene en el siguiente orgen
    '--------------------------------------------*
    ' elemento  |  contenido                    |*
    '   0       |   Codigo del Inmueble         |*
    '   1       |   Codigo del propietario      |*
    '   2       |   Nombre del Propietario      |*
    '   3       |   inicial                     |*
    '   4       |   deuda financiada            |*
    '   5       |   Nº cuotas                   |*
    '   6       |   Estatus del convenio        |*
    '--------------------------------------------*
    If datos(0) = "" Then
        strCadena = "Revise el código del inmueble en la ficha datos generales"
    ElseIf datos(1) = "" Then
        strCadena = "Revise el còdigo del apartamento en la ficha datos generales"
    ElseIf datos(2) = "" Then
        strCadena = "Revise el nombre del propietario en la ficha datos generales"
    ElseIf datos(3) = "" Or datos(3) < 1 Then
        strCadena = "Revise el monto de la inicial"
    ElseIf datos(4) = "" Or datos(4) < 1 Then
        strCadena = "Revise el monto a financiar"
    ElseIf datos(5) = "" Or datos(5) = 0 Then
        strCadena = "Falta especificar el Nº de cuotas"
    ElseIf datos(6) = "" Then
        strCadena = "Falta el Estatus del convenio. Pongase en contacto con el administrador " _
        & "del sistema"
    End If
    If strCadena <> "" Then Datos_Necesarios = MsgBox(strCadena, vbExclamation, App.ProductName)
    
    End Function

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (X >= 7770 And X <= 8770) And (Y >= 6770 And Y <= 7070) Then
    Picture1.Visible = True
Else
    Picture1.Visible = False
End If

End Sub

    Private Sub wsServidor_ConnectionRequest(ByVal requestID As Long)
    Dim Frm As New frmMsg
    Call Frm.wsServidor_ConnectionRequest(requestID)
    Frm.dtc.Enabled = False
    Frm.cmd(0).Enabled = True
    Frm.Timer1(0).Interval = 10
    Load Frm
    End Sub

    Private Sub wsServidor_Error(ByVal Number As Integer, Description As String, ByVal Scode _
    As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, _
    CancelDisplay As Boolean)
    wsServidor.Close
    MsgBox Description & vbCrLf & "Se ha detenido el servicio de Mensajería SAC", vbCritical, App.ProductName
    End Sub
