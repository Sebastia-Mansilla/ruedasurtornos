from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = r"D:\opencode\mipre"

def money(v):
    return "$" + format(int(round(v)), ",").replace(",", ".")

foda = [
    ("Fortalezas", "Diseño, electrónica y soldadura propias; taller propio en Villa Mercedes; crédito MiPre de $3.000.000 a sola firma; dos modelos diferenciados con garantía y postventa local; cercanía con tornería y herrería para mecanizado."),
    ("Debilidades", "Marca nueva sin reconocimiento; capital propio limitado a $500.000; mecanizado de ejes y platos tercerizado; dependencia de proveedores de motores; sin historial de ventas; responsable inscripto con mayor carga administrativa."),
    ("Oportunidades", "Demanda insatisfecha en el interior por equipos con servicio real; Hornos Lujo como derivación de mantenimiento/reparación y talleres de Villa Mercedes como clientes y aliados; programas provinciales de financiamiento; auge de la cerámica artesanal y terapéutica; importados sin asistencia técnica."),
    ("Amenazas", "Tornos Fidel ($495.000-$1.090.000) y Alfarero.com.ar con marca instalada; torno importado de 350W a $457.416,96; inflación y tipo de cambio; pérdida de poder adquisitivo; tope de $716.840,77 del monotributo; coincidencia de marca con Neumáticos RuedaSur."),
]

estrategias = [
    ("FO", "Usar las capacidades propias y el crédito para fabricar tornos premium con entrega, garantía de 12 meses y servicio en 48 h, apalancando derivaciones de mantenimiento/reparación de hornos y alianzas con talleres locales."),
    ("DO", "Contratar cupos fijos de mecanizado con la tornería local, capacitarse en costos y gestión, registrar la marca en INPI y presupuestar en dólares los insumos importados."),
    ("FA", "Diferenciarse del importado barato por servicio local, garantía y alquiler por turnos como alternativa de bajo ticket; evitar competir solo por precio."),
    ("DA", "Escalonar precios por modelo, mantener stock mínimo de motores y repuestos, constituir un fondo de contingencia y realizar el estudio de marca antes de imprimir rótulos."),
]

macro = [
    ("Político-legal", "MiPre ofrece crédito de $3.000.000 a sola firma, 17% anual, 42 meses con 6 de gracia; monotributo tiene precio unitario máximo de $716.840,77 para venta de cosas muebles.", "Oportunidad y amenaza", "Tomar responsable inscripto; inscribir actividad; habilitación municipal; estudio de marca en INPI."),
    ("Económico", "Inflación y actualizaciones de monotributo (16,8% desde agosto 2026); importados a $457.416,96; poder adquisitivo variable.", "Amenaza", "Indexar precios, cotizar insumos en moneda dura y mantener capital de trabajo de $500.000."),
    ("Demográfico", "Villa Mercedes y localidades aledañas concentran talleres, estudiantes y ceramistas que buscan salida laboral e ingresos extra.", "Oportunidad", "Segmentar por zona y llegar a UPrO, talleres Cmte. Videla 279, Sarmiento 267 y Munay Ki."),
    ("Tecnológico", "Motores monofásicos 3/4 HP y 1 HP, controladores digitales, memorias, reversa y pedal; IA para contenidos y gestión.", "Oportunidad", "Incorporar controlador digital en el modelo Taller y usar IA para catálogo, respuestas y CRM."),
    ("Sociocultural", "Creciente interés por lo artesanal, la cerámica terapéutica, los productos locales y el aprendizaje en talleres.", "Oportunidad", "Combinar venta, alquiler y clases; comunica el relato 'Barro, acero y territorio'."),
    ("Natural", "Consumo eléctrico, agua para limpieza, generación de residuos de arcilla y esmaltes; exigencia de eficiencia y cuidado ambiental.", "Oportunidad y amenaza", "Motores eficientes, recuperación de arcilla, separación de residuos y ficha de consumo por torno."),
]

barreras = [
    ("Economías de escala", "Media", "Fabricantes de Buenos Aires compran motores y chapa en volumen; acá se compra en lotes chicos.", "Negociar precios por volumen con proveedores locales y agrupar compras de 3 a 5 unidades."),
    ("Diferenciación", "Alta", "Alfarero.com.ar y Tornos Fidel tienen marca y cartera; hay que generar confianza desde cero.", "Garantía escrita, servicio en 48 h, prueba gratuita en taller propio y videos de casos reales."),
    ("Necesidad de capital", "Media-alta", "Se requieren herramientas, stock y capital de trabajo; el crédito MiPre cubre la barrera.", "Usar $3.000.000 del crédito + $500.000 propios; arrancar con dos modelos solamente."),
    ("Canales de distribución", "Media", "Canales ocupados por marcas conocidas; ferias y redes son de acceso abierto.", "Venta directa, Instagram/WhatsApp, ferias UPrO, consignación en talleres y derivación de mantenimiento/reparación con Hornos Lujo, sin fabricación de hornos."),
    ("Requisitos legales", "Media", "RI, habilitación municipal, facturación y normas de seguridad eléctrica.", "Contador, inscripción provincial y municipal, fichas técnicas y manual de seguridad por equipo."),
]

proveedores = [
    ("IndustryStore / MEC / Czerweny", "Motores monofásicos 3/4 HP y 1 HP", "Muy importante", "Por calidad, disponibilidad y precio"),
    ("Bulonfer (DAF)", "Motor 1 HP 1500 rpm $464.640; 3/4 HP $423.615", "Muy importante", "Por precio y origen nacional"),
    ("CORAM / Baires Rodamientos", "Rodamientos y retenes", "Importante", "Por calidad y cobertura nacional"),
    ("Ferreterías y aceros locales", "Chapa, tubos, perfiles y pintura", "Importante", "Por cercanía y forma de pago"),
    ("Tornería mecánica Urquiza 266", "Mecanizado de eje, plato y contrapuntos", "Muy importante", "Por cercanía y capacidad técnica"),
    ("Herrería metálica Colombia 320", "Corte y soldadura de estructura", "Importante", "Por cercanía y disponibilidad"),
    ("Proveedores de electrónica", "Controlador, pedal, disyuntor y cableado", "Importante", "Por precio y disponibilidad"),
]

competidores = [
    ("Tornos Fidel SRL", "Mini $495.000; pedal $680.000; con patas $890.000; con estecas $1.090.000", "Fabricante nacional con catálogo amplio", "Local / online", "Precio y variedad"),
    ("Alfarero.com.ar", "Tornos de pie y portátiles; accesorios", "Préstamo gratis en talleres y envíos al interior", "CABA / Italia", "Tecnología y prueba gratuita"),
    ("Tecnoalfarero", "Torno profesional M11, mesa M16, shablon MS-22", "Equipamiento para alfarería", "Online / Buenos Aires", "Catálogo técnico"),
    ("Torno importado 350 W", "$457.416,96 con 20% off; plato 25 cm, 4 kg, 0-350 rpm", "Bajo precio, sin servicio local", "Importado", "Precio"),
    ("VEVOR / plataformas", "Mini tornos y kits de alfarería", "Precio bajo y entrega rápida", "Importado", "Precio y volumen"),
]

sustitutos = [
    ("Torno manual o de pedal", "Bajo costo y sin electricidad", "Menor productividad y esfuerzo físico", "Artesanos ocasionales"),
    ("Alquiler de turnos en talleres", "Sin inversión inicial", "Disponibilidad limitada de horarios", "Principiantes y ceramistas independientes"),
    ("Servicio de horneadas y clases", "Resuelve la necesidad sin comprar equipo", "No permite producir en casa", "Hobby y terapia"),
    ("Extrusoras y modelado manual", "Otras técnicas de forma", "No sirve para piezas redondas en torno", "Producción alternativa"),
    ("Torno usado o restaurado", "Precio más bajo", "Sin garantía ni servicio", "Compradores con restricción presupuestaria"),
]

roles = [
    ("Iniciador", "Ceramista o docente que detecta que su torno viejo falla o que necesita ampliar el taller.", "Visita al taller, ferias, recomendación de otro ceramista."),
    ("Prescriptor", "Docente de UPrO, ceramista reconocido o taller que recomienda la marca.", "Prueba gratuita, videos, testimonios y fichas técnicas."),
    ("Decisor", "Dueño del taller o ceramista profesional que elige modelo y presupuesto.", "Comparación de fichas, garantía y servicio local."),
    ("Comprador", "Quien paga: el taller, el socio o un familiar que regala el equipo.", "Factura A, medios de pago, entrega y financiación."),
    ("Consumidor", "Alumno o ceramista que usa el torno y evalúa estabilidad, ruido y pedal.", "Experiencia de uso, recomendación y renovación."),
]

segmentacion = [
    ("Geográfica", "Villa Mercedes, San Luis Capital, Juana Koslay, La Punta y locales del interior con dificultad para servicio técnico.", "Cercanía, entrega propia y respuesta rápida."),
    ("Demográfica", "Ceramistas 25-55 años, docentes, estudiantes de arte/Oficios, talleres B2B con 1-5 tornos.", "Poder de compra medio-alto para el modelo Taller."),
    ("Psicográfica", "Personas que valoran lo artesanal, la terapia con barro, la estética y el orgullo por lo local.", "Relato de marca y diseño con identidad puntana."),
    ("Conductual", "Compra por necesidad de reemplazo, ampliación de taller, inicio de actividad o regalo; usa el torno semanalmente.", "Bono de alquiler, mantenimiento preventivo y capacitación."),
]

personas = [
    ("Ceramista independiente", "30-45 años, vende piezas en ferias e Instagram; necesita estabilidad y pedal sensible.", "Modelo Estudio + plan de mantenimiento."),
    ("Docente / taller B2B", "35-55 años, 1-5 tornos, 6-20 alumnos; prioriza durabilidad y soporte.", "Modelo Taller, alquiler de turnos y descuento por 2 unidades."),
    ("Estudiante / terapia", "18-35 años, quiere probar sin comprar; presupuesto bajo por turno.", "Alquiler por turno y clases mensuales."),
]

modelos = [
    ("Rueda Sur Estudio", "1/2 HP monofásico 220 V", "0-300 rpm con pedal", "28 cm aluminio", "Aprox. 8 kg", "Batea plástica, estructura tubular, garantía 12 meses", money(1690000)),
    ("Rueda Sur Taller", "1 HP monofásico 220 V", "0-350 rpm, reversa y memorias", "35 cm aluminio reforzado", "Aprox. 25 kg", "Batea reforzada, control digital, garantía 12 meses", money(2790000)),
]

costos = [
    ("Motor monofásico", 396125, 464640),
    ("Estructura, chapa y tubos", 95000, 150000),
    ("Plato, eje y rodamientos (mecanizado tercerizado)", 130000, 220000),
    ("Batea y separadores", 45000, 80000),
    ("Pedal, controlador y cableado", 85000, 190000),
    ("Pintura y embalaje", 30000, 50000),
    ("Mano de obra directa", 120000, 192000),
    ("Indirectos de fabricación (15%)", 135169, 201996),
    ("COSTO TOTAL", 1036294, 1548636),
    ("PRECIO DE VENTA", 1690000, 2790000),
    ("Margen bruto", 653706, 1241364),
]

costos_servicios = [
    ("Alquiler turno 3 h", money(45000), money(12000), "Luz, pasta, limpieza y desgaste"),
    ("Bono 4 turnos / mensual", money(160000), money(48000), "Cuatro turnos de 3 h"),
    ("Clase mensual de torno", money(85000), money(20000), "Pasta, esmaltes y horneada"),
    ("Visita de mantenimiento", money(70000), money(35000), "Repuestos promedio"),
]

inversion = [
    ("Torno Rueda Sur Estudio para taller propio", 1040000, "Fabricación propia con insumos del crédito"),
    ("Herramientas y equipamiento de taller (detalle en sección 9.3)", 600000, "Soldadora, taladro, mechas, amoladoras, discos, protección e instrumental"),
    ("Bancadas y mobiliario", 180000, "Instalación del taller"),
    ("Computadora y software CAD", 350000, "Diseño y electrónica"),
    ("Mejoras edilicias", 300000, "Instalación eléctrica y sectores de trabajo"),
    ("Stock inicial de insumos", 550000, "Motores, chapa, electrónica y rodamientos"),
    ("Capital de trabajo", 483706, "Aporte propio de $500.000"),
    ("TOTAL", 3500000, "Crédito MiPre $3.000.000 + capital propio $500.000"),
]

flujo_rows = []
acumulado = 0
for m in range(1, 19):
    ventas = 0 if m < 4 else (2790000 + 1690000 if m % 2 == 1 else 1690000)
    cogs = 0 if m < 4 else (2590000 if m % 2 == 1 else 1036294)
    alquiler = 160000
    clases = 0 if m == 1 else 510000
    mantenimiento = 0 if m == 1 else 210000
    total_ing = ventas + alquiler + clases + mantenimiento
    variables = cogs + 48000 + (0 if m == 1 else 120000) + (0 if m == 1 else 70000)
    fijos = 1250000
    resultado = total_ing - variables - fijos
    acumulado += resultado
    flujo_rows.append([m, ventas, alquiler, clases, mantenimiento, total_ing, variables, fijos, resultado, acumulado])

cronograma = [
    ("Nov 2026", "Cursar módulo 2 (fecha máxima 30/11 según cartilla) y definir marca, modelos y proveedores."),
    ("Dic 2026", "Módulo 3: producto, producción, costos e inversión; cerrar costos y presupuestos."),
    ("Ene 2027", "Módulo 4: marketing e IA; construir identidad, perfiles y plan de contenidos."),
    ("Feb-Mar 2027", "Presentación formal del plan y la documentación al mentor."),
    ("Mar 2027", "Aprobación, firma del crédito y compra de herramientas y stock."),
    ("Abr 2027", "Armado del torno propio, instalación del taller y primeras pruebas."),
    ("May 2027", "Inicio de alquiler, clases y mantenimiento; primeras ventas."),
    ("Jun-Sep 2027", "Venta recurrente, derivaciones de mantenimiento/reparación de hornos y alianzas con talleres, registro de marca."),
    ("Oct 2027-Mar 2028", "Segundo modelo, lanzamiento de repuestos y accesorios, evaluación del plan."),
]

riesgos = [
    ("Conflicto de marca con Neumáticos RuedaSur", "Media", "Alto", "Usar siempre 'Rueda Sur Tornos' separado y con descriptor; estudio de marca en INPI clase 7 antes de imprimir."),
    ("Dependencia de proveedores de motores", "Media", "Alto", "Dos proveedores alternativos y stock mínimo de dos unidades."),
    ("Mecanizado tercerizado", "Media", "Medio", "Acuerdo de cupos con la tornería local y plano técnico listo para tercerizar."),
    ("Inflación y suba de insumos", "Alta", "Alto", "Actualización mensual de precios, cotizaciones vigentes y capital de trabajo."),
    ("Baja adhesión de talleres B2B", "Baja", "Medio", "Prueba gratuita, garantía escrita y alquiler como puerta de entrada."),
    ("Exclusión del monotributo por precio unitario", "Alta", "Alto", "Operar como responsable inscripto y revisar esquema con contador."),
    ("Demora en aprobación del crédito", "Media", "Medio", "Plan de arranque reducido con capital propio mientras se aprueba."),
]

fuentes = [
    ("Cartilla MiPre-MiPro 3° edición", "Documento local CartillaMod1_Mipre_Mipro (1).pdf"),
    ("Módulo 2 - FODA y Macroentorno", "Documento local Módulo 2 - ANÁLISIS FODA Y MACRO ENTORNO.pdf"),
    ("Módulo 3 - Microentorno y Segmentación", "Documento local Módulo 3 - ANÁLISIS DEL MICRO ENTORNO Y SEGMENTACION.pdf"),
    ("Programa oficial Mi Próximo Paso", "https://sanluis.gov.ar/desarrollo-productivo/proximo-paso/"),
    ("Créditos, beca y cronograma 3° edición", "https://agenciasanluis.com/2026/10/02/1217031-todo-lo-que-hay-que-saber-sobre-la-tercera-edicion-de-mi-proximo-paso"),
    ("Beca pagada en plataforma Sanluiseña", "https://agenciasanluis.com/2026/11/11/1230657-beneficiarios-de-la-beca-de-mi-primer-emprendimiento-cobraran-a-traves-de-la-plataforma-sanluisena"),
    ("Categorías y topes del monotributo ARCA", "https://www.afip.gob.ar/monotributo/categorias.asp"),
    ("Tornos Fidel SRL y precios", "https://marketplace.wishopai.com/vendedor/tornos-fidel-srl?lang=es"),
    ("Torno importado 350 W", "https://dpcolors.com/productos/torno-alfarero-cerrado-importado-mecanico-crev"),
    ("Fabricante Alfarero Tornos", "https://www.alfarero.com.ar/"),
    ("Tecnoalfarero", "https://tecnoalfarero.com.ar/"),
    ("Retorneador de Cerámica", "https://www.retorneadordeceramica.com.ar/"),
    ("Hornos Lujo en Villa Mercedes", "https://www.cadena3.com/noticia/la-argentina-posible/desde-villa-mercedes-al-pais-fabrican-hornos-de-ceramica-para-todo-el-interior_495368"),
    ("Precios de motores monofásicos", "https://www.industrystore.com.ar/motores-electricos-monofasicos?orderby=15"),
    ("Motor DAF 3/4 HP y 1 HP", "https://www.bulonfer.com/motor-monofasico-1500rpm-3-4hp-daf/p"),
    ("Rodamientos CORAM", "https://www.coram.com.ar/"),
    ("Referencia de precios de clases de torno", "https://artesanadelbarro.com.ar/"),
    ("Neumáticos RuedaSur (conflicto de marca)", "https://neumaticosruedasur.com.ar/"),
]

bom_rows = [
    ("Estudio", "Motor monofásico 1/2 HP 220 V", 1, "u", 396125, 396125, "IndustryStore / MEC / Bulonfer", "Crítico", "Motor DAF 1/2 HP", "Material"),
    ("Estudio", "Chapa 1,5 mm para base y batea", 1, "set", 35000, 35000, "Aceros y ferreterías locales", "Alta", "Chapa 2 mm si hay stock", "Material"),
    ("Estudio", "Tubo rectangular 40x40 y 30x30", 1, "set 6 m", 40000, 40000, "Aceros y ferreterías locales", "Alta", "Tubo 30x30 en estructura", "Material"),
    ("Estudio", "Perfiles ángulo y soportes", 1, "set", 20000, 20000, "Aceros y ferreterías locales", "Media", "Soportes soldados", "Material"),
    ("Estudio", "Plato de aluminio 28 cm", 1, "u", 55000, 55000, "Aluminio local / tornería", "Crítico", "Plato de 30 cm con ajuste", "Material"),
    ("Estudio", "Eje de acero Ø20 mm mecanizado", 1, "u", 30000, 30000, "Tornería Urquiza 266", "Crítico", "Eje Ø22 mm", "Servicio"),
    ("Estudio", "Rodamientos y retenes", 1, "set", 25000, 25000, "CORAM / Baires Rodamientos", "Crítico", "Rodamiento cerrado equivalente", "Material"),
    ("Estudio", "Mecanizado y balanceo de plato", 1, "servicio", 20000, 20000, "Tornería Urquiza 266", "Alta", "Balanceo manual", "Servicio"),
    ("Estudio", "Batea plástica 60 L", 1, "u", 25000, 25000, "Ferretería / proveedor plásticos", "Media", "Batea 50 L", "Material"),
    ("Estudio", "Separadores, protectores y tornillería", 1, "set", 20000, 20000, "Ferretería local", "Media", "Kit de tornillería surtida", "Material"),
    ("Estudio", "Pedal con potenciómetro", 1, "u", 30000, 30000, "Proveedor de electrónica", "Crítico", "Pedal analógico genérico", "Material"),
    ("Estudio", "Controlador analógico 0-300 rpm", 1, "u", 40000, 40000, "Proveedor de electrónica", "Crítico", "Controlador 0-350 rpm", "Material"),
    ("Estudio", "Cableado, disyuntor y puesta a tierra", 1, "set", 15000, 15000, "Proveedor de electrónica", "Crítico", "Cableado normalizado", "Material"),
    ("Estudio", "Pintura epoxi/antiox y solventes", 1, "set", 15000, 15000, "Ferretería / pinturería", "Media", "Pintura sintética endurecida", "Material"),
    ("Estudio", "Embalaje, espuma y etiquetas", 1, "set", 15000, 15000, "Papelería / imprenta local", "Baja", "Cartón reciclado reforzado", "Material"),
    ("Estudio", "Mano de obra directa", 8, "h", 15000, 120000, "Interno", "Alta", "Armado en dos jornadas", "Costo interno"),
    ("Estudio", "Indirectos de fabricación (15%)", 1, "calculado", 135169, 135169, "Interno", "Media", "Revisar al cierre de cada unidad", "Costo interno"),
    ("Taller", "Motor monofásico 1 HP 220 V", 1, "u", 464640, 464640, "IndustryStore / MEC / Bulonfer", "Crítico", "Motor DAF 1 HP", "Material"),
    ("Taller", "Chapa 2 mm reforzada", 1, "set", 70000, 70000, "Aceros y ferreterías locales", "Alta", "Chapa 1,5 mm doblada", "Material"),
    ("Taller", "Tubo rectangular 50x50 y 40x40", 1, "set 8 m", 60000, 60000, "Aceros y ferreterías locales", "Alta", "Tubo 40x40 reforzado", "Material"),
    ("Taller", "Perfiles ángulo y soportes", 1, "set", 20000, 20000, "Aceros y ferreterías locales", "Media", "Soportes soldados", "Material"),
    ("Taller", "Plato de aluminio reforzado 35 cm", 1, "u", 100000, 100000, "Aluminio local / tornería", "Crítico", "Plato 32 cm balanceado", "Material"),
    ("Taller", "Eje de acero Ø25 mm mecanizado", 1, "u", 50000, 50000, "Tornería Urquiza 266", "Crítico", "Eje Ø30 mm", "Servicio"),
    ("Taller", "Rodamientos reforzados y retenes", 1, "set", 40000, 40000, "CORAM / Baires Rodamientos", "Crítico", "Rodamiento sellado equivalente", "Material"),
    ("Taller", "Mecanizado, balanceo y contrapuntos", 1, "servicio", 30000, 30000, "Tornería Urquiza 266", "Crítico", "Contrapunto estándar", "Servicio"),
    ("Taller", "Batea reforzada 80 L", 1, "u", 55000, 55000, "Ferretería / proveedor plásticos", "Media", "Batea 60 L + refuerzo", "Material"),
    ("Taller", "Separadores, protectores y tornillería", 1, "set", 25000, 25000, "Ferretería local", "Media", "Kit de tornillería Inox", "Material"),
    ("Taller", "Pedal profesional", 1, "u", 60000, 60000, "Proveedor de electrónica", "Crítico", "Pedal analógico reforzado", "Material"),
    ("Taller", "Controlador digital con memorias y reversa", 1, "u", 100000, 100000, "Proveedor de electrónica", "Crítico", "Controlador analógico + reversa", "Material"),
    ("Taller", "Cableado, disyuntor, contactor y tierra", 1, "set", 30000, 30000, "Proveedor de electrónica", "Crítico", "Cableado normalizado", "Material"),
    ("Taller", "Pintura y tratamiento anticorrosivo", 1, "set", 30000, 30000, "Pinturería local", "Media", "Pintura epoxi dos componentes", "Material"),
    ("Taller", "Embalaje reforzado", 1, "set", 20000, 20000, "Papelería / imprenta local", "Baja", "Cartón + madera liviana", "Material"),
    ("Taller", "Mano de obra directa", 12, "h", 16000, 192000, "Interno", "Alta", "Armado en tres jornadas", "Costo interno"),
    ("Taller", "Indirectos de fabricación (15%)", 1, "calculado", 201996, 201996, "Interno", "Media", "Revisar al cierre de cada unidad", "Costo interno"),
]

plan_rows = [
    (1, "Ficha técnica, plano CAD y lista de corte", 1, 2, 2, "Diseño/Electrónica", "Crédito aprobado", "No"),
    (2, "Cotización y compra de materiales", 1, 3, 3, "Compras", "BOM definido", "Sí"),
    (3, "Recepción y control de calidad de insumos", 4, 4, 1, "Producción", "Compra confirmada", "Sí"),
    (4, "Corte de chapa y tubos", 5, 6, 2, "Soldadura", "Plano y materiales listos", "Sí"),
    (5, "Soldadura de estructura", 7, 8, 2, "Soldadura", "Corte conformado", "Sí"),
    (6, "Mecanizado tercerizado de eje, plato y contrapuntos", 9, 12, 4, "Tornería local", "Plano y piezas entregadas", "Sí"),
    (7, "Armado de controlador, pedal y cableado", 10, 12, 3, "Electrónica", "Insumos eléctricos listos", "No"),
    (8, "Pintura, secado y etiquetado de estructura", 13, 14, 2, "Pintura", "Soldadura conforme", "No"),
    (9, "Ensamblaje de plato, batea y transmisión", 15, 16, 2, "Armado", "Mecanizado + pintura + electrónica", "Sí"),
    (10, "Prueba de funcionamiento de 8 horas y ajustes", 17, 18, 2, "Control de calidad", "Ensamblaje completo", "Sí"),
    (11, "Control final, número de serie y manual", 19, 19, 1, "Control de calidad", "Prueba aprobada", "Sí"),
    (12, "Embalaje y entrega", 20, 20, 1, "Logística", "Control final aprobado", "Sí"),
]

herramientas_rows = [
    ("Máquina de soldar inverter 160-200 A MMA", 1, "u", 140000, 140000, "Ferretería industrial / Bulonfer", "Crítica", "Electrodos y careta se reponen por indirectos"),
    ("Taladro percutor 850 W con maletín", 1, "u", 110000, 110000, "Ferretería local", "Crítica", "Mechas se reponen por indirectos"),
    ("Set de mechas HSS, madera y concreto 1-13 mm", 1, "set", 35000, 35000, "Ferretería local", "Alta", "Reposición trimestral"),
    ("Amoladora angular 4,5 pulgadas 850 W", 2, "u", 80000, 160000, "Ferretería industrial", "Crítica", "Discos se reponen por indirectos"),
    ("Discos de corte, desbaste y accesorios", 1, "set", 45000, 45000, "Ferretería local", "Alta", "Reposición mensual"),
    ("Careta, guantes, delantal y protección", 1, "set", 30000, 30000, "Ferretería / seguridad", "Crítica", "Reposición anual"),
    ("Herramienta manual y medición (llaves, alicates, destornilladores, calibre, nivel)", 1, "set", 80000, 80000, "Ferretería local", "Media", "Reposición según desgaste"),
]

doc = Document()
section = doc.sections[0]
section.top_margin = Cm(2)
section.bottom_margin = Cm(2)
section.left_margin = Cm(2.2)
section.right_margin = Cm(2.2)
styles = doc.styles
styles["Normal"].font.name = "Calibri"
styles["Normal"].font.size = Pt(10)
styles["Normal"].paragraph_format.space_after = Pt(6)
for name, size in [("Title", 24), ("Heading 1", 15), ("Heading 2", 12)]:
    styles[name].font.name = "Calibri"
    styles[name].font.size = Pt(size)
    styles[name].font.color.rgb = RGBColor(31, 78, 121)

def shade(cell, color):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)

def add_table(headers, rows, widths=None, font_size=8):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = 1
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = str(h)
        shade(c, "D9E2F3")
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(font_size)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
            for p in cells[i].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(font_size)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return t

def bullets(items):
    for it in items:
        doc.add_paragraph(it, style="List Bullet")

def para(text, bold=False):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    return p

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("RUEDA SUR TORNOS")
r.bold = True
r.font.size = Pt(30)
r.font.color.rgb = RGBColor(31, 78, 121)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Tornos y herramientas para cerámica")
r.font.size = Pt(14)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Barro, acero y territorio")
r.italic = True
r.font.size = Pt(12)
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Plan de Negocio · Mi Primer Emprendimiento (MiPre)\nMi Próximo Paso · 3° edición\nVilla Mercedes, San Luis · Noviembre 2026")
r.font.size = Pt(11)
doc.add_page_break()

doc.add_heading("Índice", level=1)
for i, s in enumerate([
    "Resumen ejecutivo", "Ficha del emprendimiento", "Programa, etapas y cronograma",
    "Propuesta de valor y líneas de ingreso", "Análisis FODA", "Análisis del macroentorno",
    "Análisis del microentorno", "Conociendo a los clientes y segmentación",
    "Diseño del producto y proceso productivo", "Costos y precios",
    "Plan de inversión, financiamiento y flujo de caja", "Marketing e inteligencia artificial",
    "Cronograma de implementación, riesgos y checklist", "Fuentes y supuestos",
], 1):
    doc.add_paragraph(f"{i}. {s}")
doc.add_page_break()

doc.add_heading("1. Resumen ejecutivo", level=1)
para("Rueda Sur Tornos es un emprendimiento en etapa inicial radicado en Villa Mercedes, San Luis, dedicado al diseño y la fabricación de tornos alfareros eléctricos de gama premium local, su alquiler por turnos, la explotación de un taller propio y el servicio técnico de mantenimiento y repuestos. La sociedad está integrada por dos socios mayores de 18 años con residencia mínima de dos años en la provincia y se postula a la línea Mi Primer Emprendimiento (MiPre) del programa Mi Próximo Paso, 3° edición. No se solicita la beca de capacitación porque ambos socios tienen ingresos formales.")
para("La propuesta solicita un crédito de $3.000.000 a sola firma, con tasa del 17% anual, 42 meses de plazo y 6 meses de gracia, destinado a maquinaria, herramientas, tecnología y mejoras edilicias, complementado con $500.000 de capital propio. La operación se estructura como responsable inscripto para poder vender tornos por encima del precio unitario máximo del monotributo ($716.840,77). El equipo cuenta con capacidades propias de diseño, electrónica, soldadura y electricidad, y terceriza el mecanizado de ejes y platos en la tornería local de la calle Urquiza.")
para("El catálogo inicial tiene dos modelos —Rueda Sur Estudio y Rueda Sur Taller— y diez líneas de ingreso posibles, con foco inmediato en venta, alquiler, taller propio y mantenimiento. Se incorpora una lista de materiales por modelo y un plan de trabajo con diagrama de Gantt de 20 días hábiles para la primera unidad. El punto de equilibrio mensual se estima en el mes 4 de operación y el flujo acumulado se vuelve positivo en el mes 7. El riesgo principal es la coincidencia del nombre con Neumáticos RuedaSur, por lo que se usa siempre la denominación completa y separada y se prevé estudio de marca en INPI antes de imprimir rótulos.")

doc.add_heading("2. Ficha del emprendimiento", level=1)
add_table(["Campo", "Contenido"], [
    ("Nombre de marca", "Rueda Sur Tornos"),
    ("Descriptor", "Tornos y herramientas para cerámica"),
    ("Eslogan", "Barro, acero y territorio"),
    ("Rubro", "Fabricación de máquinas para alfarería, alquiler, taller propio y servicio técnico"),
    ("Localización", "Villa Mercedes, departamento General Pedernera, San Luis"),
    ("Socios", "Dos socios, mayores de 18 años, con dos años de residencia en la provincia (completar datos personales)"),
    ("Programa", "MiPre - Mi Primer Emprendimiento, 3° edición"),
    ("Beca", "No aplica: ambos socios tienen ingreso formal mensual"),
    ("Crédito solicitado", money(3000000) + " a sola firma, 17% anual, 42 meses, 6 de gracia"),
    ("Capital propio", money(500000)),
    ("Régimen tributario", "Responsable inscripto, con facturación A y recuperación de IVA"),
    ("Modelos", "Rueda Sur Estudio y Rueda Sur Taller"),
    ("Capacidades internas", "Diseño CAD, electrónica, soldadura e instalación eléctrica"),
    ("Tercerizaciones", "Mecanizado de eje, plato y contrapuntos; cortes especiales de chapa"),
])

doc.add_heading("3. Programa, etapas y cronograma", level=1)
para("El programa Mi Próximo Paso, 3° edición, organiza cuatro etapas: capacitación y formulación del plan de negocio con mentoreo; presentación formal del plan y la documentación; aprobación y financiamiento; y puesta en marcha con mentoreo de seguimiento. Para MiPre el crédito máximo es de $3.000.000 a sola firma, con garantía solidaria no exigida, y la tasa es 17% anual con 42 meses de plazo, 6 de gracia y 36 cuotas en pesos iguales. El destino del crédito incluye equipamiento, maquinaria, tecnologías, mejoras edilicias, nuevos clientes y nuevos mercados; no es compatible con compra y venta.")
add_table(["Etapa", "Fecha oficial", "Situación al 24/11/2026"], [
    ("Inscripción", "2 al 28 de octubre de 2026", "Finalizada"),
    ("Módulo 1 - Clase presencial", "31 de octubre de 2026", "Cursado"),
    ("Módulo 2 - Mercado, FODA y clientes", "Noviembre de 2026", "Cursado; cartilla fija fecha máxima 30/11"),
    ("Módulo 3 - Producto, producción, costos e inversión", "Diciembre de 2026", "Pendiente; este dossier lo anticipa"),
    ("Módulo 4 - Marketing e IA", "Enero de 2027", "Pendiente; este dossier lo anticipa"),
    ("Presentación del plan y documentación", "1 de febrero al 15 de marzo de 2027", "A preparar"),
    ("Aprobación y financiamiento", "Desde el 1 de marzo de 2027", "A seguir con el mentor"),
    ("Puesta en marcha con mentoreo", "Marzo a diciembre de 2027", "Plan operativo incluido"),
])
para("Observación: la cartilla entregada indica 'A partir 1° de febrero del 2027' para la presentación y fechas máximas por clase (30/11, 31/12 y 31/01), mientras que el cronograma oficial publicado precisa el rango 1 de febrero al 15 de marzo de 2027. Ante cualquier discrepancia prevalece lo que indique el sistema y el mentor asignado.")

doc.add_heading("4. Propuesta de valor y líneas de ingreso", level=1)
para("Fabricamos tornos pensados para talleres del interior que necesitan estabilidad, pedal sensible y una garantía con respuesta rápida. La diferenciación no es solo técnica: es la combinación de diseño propio, fabricación local, entrega directa y servicio en 48 horas. El alquiler permite que un ceramista pruebe antes de comprar y genera relación comercial desde el primer mes.")
add_table(["Salida", "Descripción", "Prioridad", "Horizonte"], [
    ("Venta de tornos", "Dos modelos premium con garantía de 12 meses", "Core", "Mes 4"),
    ("Alquiler de tornos", "Turnos de 3 h y bonos mensuales en taller propio", "Core", "Mes 1"),
    ("Taller propio y clases", "Clases mensuales de torno y producción de piezas", "Core", "Mes 2"),
    ("Mantenimiento y reparación", "Visitas técnicas, ajustes y reemplazo de rodamientos", "Core", "Mes 2"),
    ("Repuestos", "Rodamientos, pedales, correas, capacitor y platos", "Fase 2", "Mes 6"),
    ("Accesorios y herramientas", "Brazos, tornetas, cortantes y porta-herramientas", "Fase 2", "Mes 6"),
    ("Capacitación técnica", "Cursos breves de armado y mantenimiento para ceramistas", "Fase 2", "Mes 9"),
    ("Kits y modelo portátil", "Kit de armado y torno de mesa para principiantes", "Fase 3", "Mes 12"),
    ("Mantenimiento/reparación de hornos", "Servicio eléctrico y mecánico de hornos de cerámica por derivación; no fabricación ni venta de hornos", "Fase 2", "Mes 6"),
    ("Alianzas y consignación", "Acuerdo con talleres para mostrar y vender tornos; con Hornos Lujo solo derivación de mantenimiento/reparación, sin fabricación", "Fase 2", "Mes 6"),
])

doc.add_heading("5. Análisis FODA", level=1)
para("El FODA combina factores internos (fortalezas y debilidades) y externos (oportunidades y amenazas). La matriz se traduce en acciones concretas para confirmar, ajustar o reorientar el proyecto antes de ponerlo en marcha.")
add_table(["Cuadrante", "Contenido"], foda)
add_table(["Estrategia", "Acción"], estrategias)

doc.add_heading("6. Análisis del macroentorno", level=1)
para("Se analizan las seis fuerzas del entorno —política-legal, económica, demográfica, tecnológica, sociocultural y natural— para reconocer oportunidades y amenazas sobre las que el emprendimiento no tiene control directo, pero puede anticiparse.")
add_table(["Fuerza", "Hecho relevante", "Tipo", "Acción"], macro)

doc.add_heading("7. Análisis del microentorno", level=1)
doc.add_heading("7.1 Barreras de ingreso", level=2)
add_table(["Barrera", "Nivel", "Por qué afecta", "Cómo se aborda"], barreras)
doc.add_heading("7.2 Proveedores", level=2)
add_table(["Proveedor", "Ítem", "Importancia", "Criterio"], proveedores)
para("Los proveedores críticos son los de motores y el mecanizado externo. Se prevé trabajar con al menos dos fuentes de motores y un acuerdo de cupos con la tornería local para reducir el poder de negociación y evitar demoras.")
doc.add_heading("7.3 Competencia directa", level=2)
add_table(["Competidor", "Productos y precios", "Diferencial", "Alcance", "Por qué lo eligen"], competidores)
doc.add_heading("7.4 Competencia indirecta y sustitutos", level=2)
add_table(["Sustituto", "Ventaja", "Limitación", "Quién lo elige"], sustitutos)
doc.add_heading("7.5 Clientes y poder de negociación", level=2)
para("El poder negociador de los clientes es medio-alto: pueden comparar con Tornos Fidel, con importados y con la opción de alquilar. Se reduce con diferenciación, garantía escrita, prueba gratuita y servicio local, de modo que el precio no sea el único criterio de elección.")
add_table(["Rol", "Quién sería en este proyecto", "Cómo llegar"], roles)

doc.add_heading("8. Conociendo a los clientes y segmentación", level=1)
para("El mercado definido son los talleres de cerámica, ceramistas independientes, docentes y estudiantes de Villa Mercedes y localidades cercanas, con necesidades de equipamiento, reemplazo de tornos, práctica y salida laboral. Se segmenta con los cuatro criterios del módulo 3 y se distingue entre consumidor final (B2C) y empresas u organizaciones (B2B).")
add_table(["Tipo de segmentación", "Segmento identificado", "Implicancia comercial"], segmentacion)
add_table(["Buyer persona", "Características", "Oferta sugerida"], personas)
para("Los segmentos elegidos son: talleres y ceramistas exigentes para la venta del modelo Taller; ceramistas independientes para el modelo Estudio; estudiantes y público terapéutico para alquiler y clases; y estudios/talleres B2B para mantenimiento y consignación. No se intenta vender a todos: se concentran esfuerzos en grupos con mayor probabilidad de compra y cercanía geográfica.")

doc.add_heading("9. Diseño del producto y proceso productivo", level=1)
add_table(["Modelo", "Motor", "Velocidad", "Plato", "Carga", "Características", "Precio"], modelos)
para("Proceso productivo: (1) diseño en CAD y ficha técnica; (2) corte de chapa y tubos; (3) soldadura de estructura; (4) mecanizado tercerizado de eje, plato y contrapuntos; (5) pintura y secado; (6) armado del controlador, pedal y cableado; (7) montaje del plato y la batea; (8) prueba de funcionamiento de 8 horas; (9) control de calidad y etiquetado; (10) embalaje y entrega.")
para("Layout del taller: recepción y stock, sector de corte y soldadura, sector de electrónica y armado, sector de pintura, sector de ensamblaje, sector de pruebas y oficina. El espacio propio permite organizar flujo en una sola dirección y mantener el torno de demostración en el área de atención al cliente.")
para("Capacidad: con dos personas y el mecanizado externo se puede producir un torno Estudio por mes y un torno Taller cada dos meses, además del equipo destinado al taller propio. La capacidad se ajusta aumentando los cupos de mecanizado, no con más inversiones en maquinaria pesada.")
bullets([
    "Control de calidad: centrado del plato, vibración, ruido menor a 65 dB, rango de rpm, respuesta del pedal, freno, reversa y estanqueidad de la batea.",
    "Seguridad eléctrica: disyuntor, puesta a tierra, cableado normalizado, protección IP54 en el controlador y manual de usuario por equipo.",
    "Documentación: ficha técnica, plan de mantenimiento, registro de horas de prueba y número de serie; garantía escrita de 12 meses.",
])

doc.add_heading("9.1 Lista de materiales (BOM)", level=2)
para("La lista de materiales desagrega cada torno por ítem, cantidad, proveedor y criticidad. Los montos son estimaciones de noviembre de 2026 y deben reconfirmarse con presupuestos vigentes antes de cada compra. El proveedor crítico es el mecanizado externo de eje, plato y contrapuntos.")
add_table(["Modelo", "Ítem", "Cantidad", "Costo", "Proveedor/origen", "Criticidad"],
          [[r[0], r[1], f"{r[2]} {r[3]}", money(r[5]), r[6], r[7]] for r in bom_rows],
          widths=[1.8, 5.2, 1.8, 2.2, 4.2, 2.0], font_size=7)
add_table(["Modelo", "Costo directo", "Indirectos", "Costo total"], [
    [modelo,
     money(sum(r[5] for r in bom_rows if r[0] == modelo and r[9] != "Costo interno")),
     money(sum(r[5] for r in bom_rows if r[0] == modelo and r[9] == "Costo interno")),
     money(sum(r[5] for r in bom_rows if r[0] == modelo))]
    for modelo in ("Estudio", "Taller")
], widths=[3.0, 4.0, 4.0, 4.0], font_size=8)
bullets([
    "Alternativas previstas: segundo proveedor de motores, rodamientos equivalentes de CORAM/Baires y platos de aluminio con otro diámetro nominal.",
    "Los ítems críticos no deben sustituirse sin revisar ficha técnica, balanceo y prueba de 8 horas.",
])

doc.add_heading("9.2 Plan de trabajo y diagrama de Gantt", level=2)
para("El plan de trabajo contempla 20 días hábiles para la primera unidad, con tareas paralelas donde es posible. En unidades siguientes, con planos, stock y cupos de mecanizado confirmados, el plazo estimado baja a 12-15 días hábiles.")
def gantt_bar(start, end, total=20):
    return "·" * (start - 1) + "█" * (end - start + 1) + "·" * (total - end)
add_table(["#", "Tarea", "Días", "Responsable", "Depende de", "Gantt D1-D20"],
          [[tid, tarea, f"D{ini}-D{fin} ({dur} d)", resp, dep, gantt_bar(ini, fin)]
           for tid, tarea, ini, fin, dur, resp, dep, crit in plan_rows],
          widths=[0.8, 5.0, 2.2, 3.0, 4.5, 4.0], font_size=7)
bullets([
    "Ruta crítica: compra de materiales → corte y soldadura → mecanizado tercerizado → ensamblaje → prueba final → entrega.",
    "El mecanizado externo es el principal cuello de botella: confirmar cupo con la tornería antes de aceptar pedidos.",
    "El alcance con hornos se limita a mantenimiento/reparación por derivación; no se fabrican ni venden hornos.",
])

doc.add_heading("9.3 Herramientas y equipamiento de taller", level=2)
para("Las herramientas se tratan como inversión inicial, no como costo unitario de cada torno. El presupuesto previsto es de $600.000 y se financia con el crédito MiPre. Los consumibles de reposición —mechas, discos, electrodos y elementos de protección— se cubren con los indirectos de fabricación del 15%, para no duplicarlos en el costo de cada equipo.")
add_table(["Herramienta", "Cant.", "Costo", "Proveedor/origen", "Prioridad", "Reposición"],
          [[h[0], f"{h[1]} {h[2]}", money(h[4]), h[5], h[6], h[7]] for h in herramientas_rows],
          widths=[6.0, 1.8, 2.2, 4.0, 2.0, 4.0], font_size=7)
add_table(["Concepto", "Monto"], [
    ("Total herramientas y equipamiento", money(sum(h[4] for h in herramientas_rows))),
    ("Límite previsto en inversión", money(600000)),
], widths=[8.0, 5.0], font_size=8)
bullets([
    "Comprar primero soldadora, taladro, amoladoras y protección; diferir herramienta manual y medición si hay restricción de caja.",
    "Registrar cada herramienta con número de inventario, fecha de compra y vida útil estimada para reposición y depreciación interna.",
])

doc.add_heading("10. Costos y precios", level=1)
para("Los costos se construyen con cotizaciones relevadas en el mercado argentino durante noviembre de 2026. El motor de 1/2 HP se estima en $396.125 y el de 1 HP en $464.640; el mecanizado se presupuesta aparte. Los indirectos se calculan en 15% sobre materiales y mano de obra directa.")
add_table(["Ítem", "Estudio", "Taller"], [[a, money(b), money(c)] for a, b, c in costos])
para("El detalle por ítem, cantidad y proveedor se presenta en la sección 9.1 y en la planilla 'BOM_Materiales' del archivo complementario. Las herramientas de taller se detallan en la sección 9.3 y se imputan a inversión; su reposición ya está contemplada en los indirectos.")
add_table(["Servicio", "Precio", "Costo variable estimado", "Detalle"], costos_servicios)
para("Precio de venta: Rueda Sur Estudio " + money(1690000) + " y Rueda Sur Taller " + money(2790000) + ". Ambos se ubican por encima del importado de $457.416,96 y del rango de Tornos Fidel ($495.000-$1.090.000), justificados por garantía, servicio en 48 horas, diseño local y control de calidad. Esta decisión es compatible con responsable inscripto, no con monotributo.")

doc.add_heading("11. Plan de inversión, financiamiento y flujo de caja", level=1)
add_table(["Ítem", "Monto", "Destino"], [[a, money(b), c] for a, b, c in inversion])
para("Financiamiento: crédito MiPre de " + money(3000000) + " a sola firma, tasa 17% anual, 42 meses con 6 meses de gracia, y aporte propio de " + money(500000) + ". La beca de $250.000 mensuales no se solicita porque ambos socios perciben ingresos formales. El crédito se destina a equipamiento, herramientas, tecnología, mejoras edilicias y stock inicial, cumpliendo la restricción del programa que impide destinarlo a compra y venta.")
para("Flujo de caja proyectado a 18 meses. Supuestos: alquiler desde el mes 1; clases y mantenimiento desde el mes 2; ventas desde el mes 4 (un Estudio por mes y un Taller cada dos meses); costos fijos mensuales de " + money(1250000) + " incluyendo retribución de socios, servicios, contador, marketing, seguros e impuestos estimados. No incluye depreciaciones ni variaciones de precio.")
add_table(["Mes", "Ventas", "Alquiler", "Clases", "Mant.", "Ingresos", "Variables", "Fijos", "Resultado", "Acumulado"],
          [[m, money(v), money(a), money(c), money(man), money(ti), money(va), money(fi), money(re), money(ac)] for m, v, a, c, man, ti, va, fi, re, ac in flujo_rows],
          widths=[1.1, 2.1, 1.7, 1.7, 1.7, 2.1, 2.1, 1.9, 2.1, 2.1], font_size=7)
para("Punto de equilibrio mensual estimado: mes 4. Flujo acumulado positivo: mes 7. El fondo de contingencia se conforma con el excedente de los meses positivos y con el capital propio remanente.")

doc.add_heading("12. Marketing e inteligencia artificial", level=1)
para("Identidad: nombre separado y completo 'Rueda Sur Tornos' para diferenciarse de Neumáticos RuedaSur; descriptor 'Tornos y herramientas para cerámica'; eslogan 'Barro, acero y territorio'. Paleta sugerida: barro, negro industrial y acero; tipografía legible y técnica; logo conceptual con un plato girando sobre una línea de horizonte puntano.")
bullets([
    "Canales propios: Instagram y Facebook con demostraciones de centrado, pedal, ruido y mantenimiento; WhatsApp para consultas y reservas; catálogo PDF actualizado.",
    "Canales locales: ferias y encuentros en UPrO, visitas a talleres Cmte. Videla 279, Emunah y Munay Ki, y derivación de mantenimiento/reparación de hornos con Hornos Lujo, sin fabricación ni venta de hornos.",
    "Canales digitales: publicaciones con IA para generar guiones de reels, respuestas frecuentes, títulos de productos y textos de catálogo, siempre revisados por el equipo.",
    "CRM: planilla de consultas con etapa (interés, prueba, presupuesto, venta, postventa), fecha de seguimiento y origen del contacto.",
    "Mantenimiento recurrente: recordatorios de revisión cada 6 meses y venta de kits de rodamientos y pedal.",
])
add_table(["KPI", "Meta estimada"], [
    ("Consultas mensuales", "80 a partir del mes 4"),
    ("Tasa de conversión a venta", "10%"),
    ("Ventas mensuales", "1 Estudio por mes y 1 Taller cada dos meses"),
    ("Ocupación del alquiler", "4 turnos por mes desde el inicio"),
    ("Alumnos de taller", "6 por mes desde el mes 2"),
    ("Visitas de mantenimiento", "2 por mes desde el mes 2"),
])

doc.add_heading("13. Cronograma de implementación, riesgos y checklist", level=1)
add_table(["Período", "Hito"], cronograma)
add_table(["Riesgo", "Probabilidad", "Impacto", "Mitigación"], riesgos)
doc.add_heading("13.1 Checklist de presentación al mentor", level=2)
bullets([
    "Plan de negocio completo con resumen ejecutivo, FODA, macroentorno, microentorno y segmentación.",
    "Fichas técnicas de los dos modelos y presupuestos de proveedores.",
    "Planilla de costos, precios, inversión y flujo de caja a 18 meses.",
    "Datos personales de los dos socios: DNI, edades, residencia acreditada en San Luis e ingresos formales.",
    "Constancias de inscripción y cursado de los cuatro módulos de capacitación.",
    "Declaración de destino del crédito acorde con equipamiento, tecnología y mejoras edilicias.",
    "Estudio de marca en INPI en curso y comprobante de dominio y redes sociales.",
])

doc.add_heading("14. Fuentes y supuestos", level=1)
para("Los datos de mercado, precios y normativa se relevaron el 24 de noviembre de 2026. Los montos de costos y flujo son estimaciones orientativas que deben confirmarse con presupuestos vigentes al momento de la presentación. La información tributaria es general y no reemplaza el asesoramiento de un contador público.")
add_table(["Fuente", "Referencia"], fuentes)
add_table(["Supuesto", "Valor"], [
    ("Fecha de referencia", "24 de noviembre de 2026"),
    ("Tipo de cambio y precios", "Pesos argentinos, precios relevados en noviembre de 2026"),
    ("Mecanizado", "Tercerizado en tornería local de Villa Mercedes"),
    ("Alquiler inicial", "4 turnos de 3 horas por mes"),
    ("Clases iniciales", "6 alumnos por mes a $85.000"),
    ("Costos fijos mensuales", money(1250000)),
    ("Garantía", "12 meses sobre tornos fabricados"),
    ("Lead time primera unidad", "20 días hábiles"),
    ("Lead time unidades siguientes", "12-15 días hábiles"),
    ("Alcance con hornos", "Solo mantenimiento/reparación, sin fabricación ni venta de hornos"),
])

footer = section.footer.paragraphs[0]
footer.text = "Rueda Sur Tornos · Plan de Negocio · MiPre 3° edición · Villa Mercedes, San Luis"
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
for r in footer.runs:
    r.font.size = Pt(8)

doc.save(OUT + r"\RuedaSur_PlanDeNegocio.docx")

wb = Workbook()
ws = wb.active
ws.title = "Resumen"
resumen = [
    ("Marca", "Rueda Sur Tornos"),
    ("Descriptor", "Tornos y herramientas para cerámica"),
    ("Eslogan", "Barro, acero y territorio"),
    ("Programa", "MiPre - 3° edición"),
    ("Ubicación", "Villa Mercedes, San Luis"),
    ("Crédito solicitado", 3000000),
    ("Capital propio", 500000),
    ("Beca", "No aplica"),
    ("Régimen", "Responsable inscripto"),
    ("Modelo 1", "Rueda Sur Estudio"),
    ("Modelo 2", "Rueda Sur Taller"),
    ("Precio Estudio", 1690000),
    ("Precio Taller", 2790000),
    ("Punto de equilibrio", "Mes 4"),
    ("Flujo acumulado positivo", "Mes 7"),
    ("Riesgo principal", "Conflicto de marca con Neumáticos RuedaSur"),
]
ws.append(["Indicador", "Valor"])
for k, v in resumen:
    ws.append([k, v])
ws["A1"].font = Font(bold=True, color="FFFFFF")
ws["B1"].font = Font(bold=True, color="FFFFFF")
for c in ws[1]:
    c.fill = PatternFill("solid", fgColor="1F4E79")
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 60
for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
    for c in row:
        if isinstance(c.value, int):
            c.number_format = '"$" #.##0'

thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def add_sheet(name, headers, rows, widths=None, money_cols=None):
    s = wb.create_sheet(name)
    s.append(headers)
    for c in s[1]:
        c.fill = PatternFill("solid", fgColor="1F4E79")
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for row in rows:
        s.append(list(row))
    for row in s.iter_rows():
        for c in row:
            c.border = border
            c.alignment = Alignment(vertical="top", wrap_text=True)
    if money_cols:
        for col in money_cols:
            for c in s.iter_cols(min_col=col, max_col=col, min_row=2):
                for cell in c:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '"$" #.##0'
    if widths:
        for i, w in enumerate(widths, 1):
            s.column_dimensions[get_column_letter(i)].width = w
    s.freeze_panes = "A2"
    return s

add_sheet("FODA", ["Cuadrante", "Contenido"], foda, [20, 120])
add_sheet("Estrategias_FODA", ["Estrategia", "Acción"], estrategias, [15, 120])
add_sheet("Macroentorno", ["Fuerza", "Hecho relevante", "Tipo", "Acción"], macro, [20, 70, 20, 70])
add_sheet("Barreras", ["Barrera", "Nivel", "Por qué afecta", "Cómo se aborda"], barreras, [25, 15, 65, 70])
add_sheet("Proveedores", ["Proveedor", "Ítem", "Importancia", "Criterio"], proveedores, [35, 45, 20, 45])
add_sheet("Competencia_Directa", ["Competidor", "Productos y precios", "Diferencial", "Alcance", "Por qué lo eligen"], competidores, [30, 55, 45, 25, 35])
add_sheet("Sustitutos", ["Sustituto", "Ventaja", "Limitación", "Quién lo elige"], sustitutos, [35, 45, 50, 35])
add_sheet("Clientes_Roles", ["Rol", "Quién sería", "Cómo llegar"], roles, [20, 70, 60])
add_sheet("Segmentacion", ["Tipo", "Segmento", "Implicancia"], segmentacion, [22, 80, 60])
add_sheet("Buyer_Personas", ["Persona", "Características", "Oferta"], personas, [28, 80, 60])
add_sheet("Modelos_Tecnicos", ["Modelo", "Motor", "Velocidad", "Plato", "Carga", "Características", "Precio"],
          [[a, b, c, d, e, f, g] for a, b, c, d, e, f, g in modelos], [25, 30, 30, 25, 18, 60, 18], [7])
add_sheet("Costos", ["Ítem", "Estudio", "Taller"], [[a, b, c] for a, b, c in costos], [55, 20, 20], [2, 3])
add_sheet("Costos_Servicios", ["Servicio", "Precio", "Costo variable", "Detalle"], costos_servicios, [30, 20, 22, 60])
add_sheet("Inversion", ["Ítem", "Monto", "Destino"], [[a, b, c] for a, b, c in inversion], [60, 20, 70], [2])
add_sheet("Flujo_Caja_18m", ["Mes", "Ventas", "Alquiler", "Clases", "Mantenimiento", "Ingresos", "Variables", "Fijos", "Resultado", "Acumulado"],
          flujo_rows, [8, 18, 16, 16, 18, 18, 18, 16, 18, 18], list(range(2, 11)))
add_sheet("Cronograma", ["Período", "Hito"], cronograma, [22, 120])
add_sheet("Riesgos", ["Riesgo", "Probabilidad", "Impacto", "Mitigación"], riesgos, [45, 18, 15, 90])
add_sheet("Fuentes", ["Fuente", "Referencia"], fuentes, [45, 110])

bom_sheet = add_sheet("BOM_Materiales", ["Modelo", "Ítem", "Cantidad", "Unidad", "Costo unitario", "Costo total", "Proveedor/origen", "Criticidad", "Alternativa", "Tipo"],
                      [list(r) for r in bom_rows], [10, 42, 10, 12, 16, 16, 32, 12, 34, 14], [5, 6])
for modelo in ("Estudio", "Taller"):
    bom_sheet.append([modelo, "TOTAL MODELO", "", "", "", sum(r[5] for r in bom_rows if r[0] == modelo), "", "", "", ""])
    for c in bom_sheet[bom_sheet.max_row]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9E2F3")
    bom_sheet.cell(row=bom_sheet.max_row, column=6).number_format = '"$" #.##0'

add_sheet("Plan_Trabajo", ["#", "Tarea", "Inicio", "Fin", "Duración", "Responsable", "Depende de", "Ruta crítica"],
          plan_rows, [5, 48, 10, 10, 12, 20, 38, 14])

g_headers = ["#", "Tarea", "Responsable"] + [f"D{i}" for i in range(1, 21)] + ["Ruta crítica"]
g_rows = [[tid, tarea, resp] + ["█" if ini <= d <= fin else "" for d in range(1, 21)] + [crit]
          for tid, tarea, ini, fin, dur, resp, dep, crit in plan_rows]
g = add_sheet("Gantt_Fabricacion", g_headers, g_rows, [5, 48, 20] + [3] * 20 + [14])
for r_idx, (tid, tarea, ini, fin, dur, resp, dep, crit) in enumerate(plan_rows, 2):
    for d in range(ini, fin + 1):
        c = g.cell(row=r_idx, column=3 + d)
        c.fill = PatternFill("solid", fgColor="F4B183" if crit == "Sí" else "9DC3E6")
        c.font = Font(bold=True)
g.cell(row=len(plan_rows) + 3, column=2, value="Leyenda: naranja = ruta crítica; azul = tarea paralela.").font = Font(italic=True)
g.freeze_panes = "D2"

herr_sheet = add_sheet("Herramientas_Taller", ["Herramienta", "Cantidad", "Unidad", "Costo unitario", "Costo total", "Proveedor/origen", "Prioridad", "Reposición"],
                       [list(h) for h in herramientas_rows], [58, 10, 10, 16, 16, 32, 12, 40], [4, 5])
herr_sheet.append(["TOTAL", "", "", "", sum(h[4] for h in herramientas_rows), "", "", ""])
for c in herr_sheet[herr_sheet.max_row]:
    c.font = Font(bold=True)
    c.fill = PatternFill("solid", fgColor="D9E2F3")
herr_sheet.cell(row=herr_sheet.max_row, column=5).number_format = '"$" #.##0'

wb.save(OUT + r"\RuedaSur_Anexos_Cuadros.xlsx")
print("OK")
