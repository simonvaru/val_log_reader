import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.append(['ID', 'Mensaje en log', 'Significado'])
rows = [
    (1,  'Checking EMV library initialization',                        'Checking EMV library initialization cada una hora'),
    (2,  'Se detectó una tarjeta EMV',                                 'Se detectó una tarjeta EMV en VL550'),
    (3,  'Tarjeta Mifare detectada',                                   'Se detectó una tarjeta Mifare en VL550'),
    (5,  '====================== INICIO VALIDACION ======================', 'Inicio de intento de cobro de pasaje'),
    (6,  '======================  FIN  VALIDACIÓN  ======================', 'Fin de intento de cobro de pasaje'),
    (7,  'COMPANY:',                                                   'Número de empresa'),
    (8,  'iAppCalEmv_Debit OK',                                        'iAppCalEmv_Debit exitoso'),
    (9,  'iAppCalEmv_Debit FAILED',                                    'iAppCalEmv_Debit falla'),
    (10, 'ulLastFour',                                                 '4 últimos números de tarjeta EMV'),
    (11, '"serial_number":',                                           'Nro serial de dispositivo validador'),
    (12, 'appName =',                                                  'Nombre de la app'),
    (13, 'appVersion =',                                               'Versión de la app'),
    (14, 'QR Record serialNumber:',                                    'QR identificador'),
    (15, '[QRInputProducer][I]return true',                            'Payload de QR: pago con QR'),
    (16, 'sMessage: Autenticacion de datos validada y correcta',       'Autenticación de datos validada y correcta'),
    (17, '[TransactionDAOBPCImpl][T]Transaction Successfully stored!', 'TRANSACCION EXITOSA después de pagar con EMV'),
    (18, 'The reader did not reply or the reply was not expected',      'El lector no respondió o respuesta inesperada'),
    (19, 'EMV library is operational',                                 'EMV library is operational'),
    (20, '[IntegrisysEMVReader][D]FARE:',                              'Tarifa aplicada'),
    (21, '[IntegrisysEMVReader][D]sExplanation: TRANSACCION EXITOSA',  'Transacción con tarjeta EMV exitosa'),
    (22, '[ConsoleInterpreter][I]t.counter:',                          'Valor de contador de boletos'),
    (23, 'Name: CONTADOR_BOLETOS, Value:',                             'Valor de contador de boletos'),
    (24, '[ApplicationProfile][I]========= HW STATUS =========',      'Indica estados de EMV, QR, Mk Eprom'),
    (25, 'params line = ',                                             'Inicio de lectura de parámetros del colectivo'),
    (26, 'merchantName:',                                              'Nombre de la empresa'),
    (27, 'driver =',                                                   'Número de legajo del chofer'),
    (28, 'Id: 1523, Name: EVENTS_NUMBER, Value:',                      'Evento (cobro de tarifa multipago) nro'),
    (29, '"versionFW":',                                               'Versión de FW de consola CGI'),
    (31, '[OCConnection][T][BKO -->]',                                 'currentStage reporte a back office'),
    (33, 'StateOpeningService()',                                       'Indica que abre el servicio'),
    (34, 'Id: 779, Name: SERVICE_ID, Value:',                          'Indica N° Servicio abierto'),
    (36, 'SET STATE 26',                                               'Estado 26: Se cierra servicio'),
    (37, 'SET STATE 23',                                               'Estado 23: Se abre servicio'),
    (38, 'SERVICE_ID',                                                 'Name: SERVICE_ID, Value: N° Servicio'),
    (39, 'Carga exitosa para archivo: RL',                             'Tabla RL id:9 - Identificador y versión'),
    (40, 'Carga exitosa para archivo: AL',                             'Tabla AL id:11 - Identificador y versión'),
    (41, 'Carga exitosa para archivo: CO',                             'Tabla CO id:3 - Identificador y versión'),
    (42, 'Carga exitosa para archivo: CD',                             'Tabla CD id:15 - Identificador y versión'),
    (43, 'Carga exitosa para archivo: LR',                             'Tabla LR id:23 - Identificador y versión'),
    (44, 'Carga exitosa para archivo: RS',                             'Tabla RS id:16 - Identificador y versión'),
    (45, 'Carga exitosa para archivo: GP',                             'Tabla GP id:1 - Identificador y versión'),
    (46, 'Carga exitosa para archivo: SG',                             'Tabla SG id:20 - Identificador y versión'),
    (47, 'Carga exitosa para archivo: LI',                             'Tabla LI id:18 - Identificador y versión'),
    (48, 'Carga exitosa para archivo: OL',                             'Tabla OL id:10 - Identificador y versión'),
]
for r in rows:
    ws.append(r)
wb.save('lista-eventos-vl550.xlsx')
print(f'xlsx guardado con {len(rows)} eventos')
