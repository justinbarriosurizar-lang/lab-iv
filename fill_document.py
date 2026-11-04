import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def style_paragraph(p, font_name="Calibri", font_size=10.5, space_after=4, space_before=0, line_spacing=1.15):
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = line_spacing

def add_styled_run(p, text, bold=False, italic=False, font_name="Calibri", font_size=10.5, color=None):
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = font_name
    run.font.size = Pt(font_size)
    if color:
        run.font.color.rgb = color
    return run

def add_code_block(cell, code_str):
    p = cell.add_paragraph()
    style_paragraph(p, font_name="Consolas", font_size=9.5, space_after=2, space_before=2, line_spacing=1.05)
    # Shading background for code
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F4F6F8')
    pPr.append(shd)
    
    add_styled_run(p, code_str, font_name="Consolas", font_size=9.5, color=RGBColor(0x24, 0x29, 0x2E))

def clear_cell(cell):
    for p in cell.paragraphs:
        p.text = ''
    # remove extra paragraphs if any
    while len(cell.paragraphs) > 1:
        p = cell.paragraphs[-1]._p
        p.getparent().remove(p)

def fill_doc():
    doc = docx.Document('lab iv.docx')

    # =========================================================================
    # Table 2: ¿Por qué los métodos de las interfaces no llevan cuerpo?
    # =========================================================================
    c = doc.tables[2].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Respuesta:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, 
        "Las interfaces en Java definen un contrato formal de comportamiento, es decir, establecen qué operaciones deben existir sin imponer cómo deben llevarse a cabo. "
        "Por convención de diseño orientado a objetos:\n"
    )
    p1 = c.add_paragraph()
    style_paragraph(p1)
    add_styled_run(p1, "1. Desacoplamiento y Abstracción: ", bold=True)
    add_styled_run(p1, "Al no contener cuerpo de implementación (métodos abstractos), se garantiza la total separación entre la especificación del servicio y su código ejecutable.\n")
    add_styled_run(p1, "2. Polimorfismo y Flexibilidad: ", bold=True)
    add_styled_run(p1, "Permite que múltiples clases implementen el mismo método de formas completamente distintas (por ejemplo, guardar en memoria volatil, persistir en base de datos relacional MySQL o enviar a un servicio en la nube), obligando a cada clase concreta a proveer su propia implementación obligatoria.\n")
    add_styled_run(p1, "3. Múltiple Herencia de Tipo: ", bold=True)
    add_styled_run(p1, "Java no permite herencia múltiple de clases para evitar conflictos de ambigüedad (problema del diamante); el uso de métodos sin cuerpo permite que una clase implemente varias interfaces con total seguridad.")

    # =========================================================================
    # Table 3: ¿Tendríamos que modificar AlmacenarTareas si cambiamos a BD?
    # =========================================================================
    c = doc.tables[3].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Respuesta:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, "No, en absoluto. No tendríamos que modificar ni una sola línea de la interfaz AlmacenarTareas.\n\n")
    add_styled_run(p0, "Justificación técnica:\n", bold=True)
    add_styled_run(p0, 
        "La interfaz define exclusivamente las firmas de las operaciones esenciales de negocio (guardar, buscarPorId, obtenerTodas, eliminar). "
        "El medio físico o tecnológico donde residen los datos (una lista en memoria RAM, un archivo JSON/XML o una base de datos PostgreSQL/MySQL) es un detalle de infraestructura e implementación.\n\n"
        "Bajo este principio, para migrar a una base de datos simplemente crearíamos una nueva clase, por ejemplo: \n"
        "public class AlmacenamientoBaseDatos implements AlmacenarTareas { ... }\n"
        "y programaríamos las consultas SQL o llamadas JDBC dentro de ella. La interfaz permanece intacta y los consumidores como TareaController no sufren ninguna alteración, cumpliendo a cabalidad con el Principio Abierto/Cerrado (OCP: abierto a extensión pero cerrado a modificación) de SOLID."
    )

    # =========================================================================
    # Table 4: Diagrama UML inicial y explicativo
    # =========================================================================
    c = doc.tables[4].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Diagrama de Clases UML del Sistema:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, "A continuación se presenta el diagrama UML que modela las clases principales del sistema (Tarea, TareaController, VistaTerminal, Main), la integración de la interfaz AlmacenarTareas y su implementación AlmacenamientoMemoria, así como la relación de las pruebas unitarias:")
    
    p_img = c.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_paragraph(p_img, space_before=6, space_after=6)
    p_img.add_run().add_picture('assets/uml_diagram.png', width=Inches(5.7))

    p_desc = c.add_paragraph()
    style_paragraph(p_desc)
    add_styled_run(p_desc, "Descripción de Componentes en el Diagrama:\n", bold=True)
    add_styled_run(p_desc, 
        "• Tarea: Modelo que encapsula los atributos de la tarea (id, titulo, descripcion, completada) con sus getters y setters.\n"
        "• AlmacenarTareas (<<interface>>): Abstracción que declara las operaciones CRUD requeridas (guardar, buscarPorId, obtenerTodas, eliminar).\n"
        "• AlmacenamientoMemoria: Clase concreta que implementa la interfaz AlmacenarTareas utilizando una lista interna (ArrayList<Tarea>).\n"
        "• TareaController: Controlador que gestiona la lógica de negocio (validar estados, IDs y transiciones), desacoplado del almacenamiento mediante inyección de dependencias.\n"
        "• VistaTerminal: Interfaz de usuario por consola que interactúa con el usuario y envía peticiones al controlador.\n"
        "• Main: Punto de entrada que ensambla la arquitectura (instancia el almacenamiento, el controlador y la vista)."
    )

    # =========================================================================
    # Table 5: Cómo integramos las dos interfaces a la solución actual
    # =========================================================================
    c = doc.tables[5].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Respuesta:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, 
        "Para integrar las dos interfaces (AlmacenarTareas y Notificador) a la solución y desacoplar los componentes se realizan las siguientes modificaciones clave:\n\n"
        "1. Integración de AlmacenarTareas:\n", bold=True
    )
    add_styled_run(p0, 
        "   • Se creó la clase AlmacenamientoMemoria que implementa AlmacenarTareas, encapsulando la lista List<Tarea>.\n"
        "   • En TareaController, se reemplazó la declaración interna directa de new ArrayList<>() por un atributo de tipo interfaz: private AlmacenarTareas almacenamiento;\n"
        "   • Se modificó el constructor de TareaController para recibir AlmacenarTareas por Inyección de Dependencias: public TareaController(AlmacenarTareas almacenamiento).\n\n"
    )
    add_styled_run(p0, "2. Integración de Notificador:\n", bold=True)
    add_styled_run(p0, 
        "   • Se puede incorporar como atributo opcional o inyectado en TareaController (private Notificador notificador).\n"
        "   • Cada vez que una tarea se crea, completa o elimina, el controlador invoca notificador.enviarMensaje(\"Tarea #\" + id + \" completada\").\n"
        "   • Esto permite crear diferentes implementaciones (NotificadorConsola, NotificadorEmail, NotificadorSMS) sin modificar el controlador.\n\n"
    )
    add_styled_run(p0, "3. Modificación en la Vista / Main:\n", bold=True)
    add_styled_run(p0, "   • En Main.java o VistaTerminal.java se inicializa la implementación concreta (new AlmacenamientoMemoria()) y se le inyecta a TareaController.")

    # =========================================================================
    # Table 6: Test en TareaControllerTest que no valida lo necesario
    # =========================================================================
    c = doc.tables[6].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Respuesta:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, 
        "El test que no valida lo necesario es: validarCambioDeEstado().\n\n"
        "¿Por qué no cumple su función?\n", bold=True
    )
    add_styled_run(p0, 
        "En el código original, el método únicamente creaba la tarea: \n"
        "Tarea t = controlador.agregarTarea(\"atender estudiantes\", \"william\");\n"
        "No contenía ninguna aserción (assert) ni llamaba a controlador.completarTarea(...) para modificar su estado. Por lo tanto, el test pasaba en verde de forma engañosa sin validar si el cambio de estado realmente ocurría.\n\n"
        "¿Qué hay que agregarle?\n", bold=True
    )
    add_styled_run(p0, 
        "1. Verificar el estado inicial: assertFalse(t.getCompletada());\n"
        "2. Ejecutar la acción bajo prueba: controlador.completarTarea(t.getId());\n"
        "3. Validar el estado final esperado: assertTrue(t.getCompletada());"
    )

    # =========================================================================
    # Table 7: Código agregado al test validarCambioDeEstado()
    # =========================================================================
    c = doc.tables[7].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Código agregado y completado en TareaControllerTest.java:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    code_table7 = (
        "@Test\n"
        "void validarCambioDeEstado() {\n"
        "    Tarea t = controlador.agregarTarea(\"atender estudiantes\", \"william\");\n"
        "    assertFalse(t.getCompletada()); // Verifica que inicialmente no está completada\n"
        "    \n"
        "    controlador.completarTarea(t.getId()); // Se ejecuta la acción de completar\n"
        "    assertTrue(t.getCompletada()); // Se verifica que su estado sea ahora true\n"
        "}"
    )
    add_code_block(c, code_table7)

    # =========================================================================
    # Table 8: Método testCompletarTareaInexistenteLanzaExcepcion()
    # =========================================================================
    c = doc.tables[8].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Método de prueba con assertThrows para ID inexistente:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    code_table8 = (
        "@Test\n"
        "void testCompletarTareaInexistenteLanzaExcepcion() {\n"
        "    // Verifica que al pasar un ID no registrado (ej. 99) lance IllegalArgumentException\n"
        "    assertThrows(IllegalArgumentException.class, () -> {\n"
        "        controlador.completarTarea(99);\n"
        "    });\n"
        "}"
    )
    add_code_block(c, code_table8)

    # =========================================================================
    # Table 9: Método testCompletarTareaYaCompletadaLanzaExcepcion()
    # =========================================================================
    c = doc.tables[9].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Método de prueba para tarea ya completada:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    code_table9 = (
        "@Test\n"
        "void testCompletarTareaYaCompletadaLanzaExcepcion() {\n"
        "    Tarea t = controlador.agregarTarea(\"Estudiar para el examen\", \"POO\");\n"
        "    controlador.completarTarea(t.getId()); // Primera vez: se completa exitosamente\n"
        "    \n"
        "    // Segunda vez: debe lanzar IllegalStateException ya que estaba completada previamente\n"
        "    assertThrows(IllegalStateException.class, () -> {\n"
        "        controlador.completarTarea(t.getId());\n"
        "    });\n"
        "}"
    )
    add_code_block(c, code_table9)

    # =========================================================================
    # Table 11: Análisis del if (id > contadorTareas) y corrección
    # =========================================================================
    c = doc.tables[11].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Análisis de la condición y corrección:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, "1. ¿Qué pasa si alguien intenta completar la tarea con ID 3?\n", bold=True)
    add_styled_run(p0, 
        "Si hay 3 tareas (con IDs 0, 1 y 2), el contadorTareas vale 3. Al evaluar if (id > contadorTareas) con id = 3, "
        "la expresión es if (3 > 3), lo cual resulta FALSO.\n"
        "Por lo tanto, NO se lanza la excepción IllegalArgumentException esperada, a pesar de que la tarea con ID 3 no existe en el sistema. "
        "El flujo continúa y genera un error silencioso o busca un elemento inexistente. Además, si el usuario pasa un ID negativo (ej. -1), la condición tampoco lo detecta.\n\n"
    )
    add_styled_run(p0, "2. ¿Cómo corregir la condición if?\n", bold=True)
    add_styled_run(p0, 
        "• Corrección basada en rangos de índice (considerando IDs 0, 1, ..., contadorTareas - 1):\n"
    )
    code_fix1 = (
        "if (id < 0 || id >= contadorTareas) {\n"
        "    throw new IllegalArgumentException(\"No hay una tarea con el ID: \" + id);\n"
        "}"
    )
    add_code_block(c, code_fix1)
    
    p_fix2 = c.add_paragraph()
    style_paragraph(p_fix2)
    add_styled_run(p_fix2, 
        "• Corrección profesional y desacoplada (utilizando el repositorio/almacenamiento):\n"
        "La forma más limpia y robusta, independiente de cómo se asignen o borren los IDs, es delegar la búsqueda al almacenamiento y validar si el objeto retornado es nulo:"
    )
    code_fix2 = (
        "Tarea t = almacenamiento.buscarPorId(id);\n"
        "if (t == null) {\n"
        "    throw new IllegalArgumentException(\"No hay una tarea con el ID: \" + id);\n"
        "}"
    )
    add_code_block(c, code_fix2)

    # =========================================================================
    # Table 12: Screenshot de los tests corrieron (Fase 1)
    # =========================================================================
    c = doc.tables[12].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Captura de pantalla: Pruebas unitarias de la primera fase ejecutándose exitosamente en verde:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    p_img = c.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_paragraph(p_img, space_before=4, space_after=4)
    p_img.add_run().add_picture('assets/screenshot_part1.png', width=Inches(5.7))

    # =========================================================================
    # Table 13: Principio SOLID que habla en contra del acoplamiento
    # =========================================================================
    c = doc.tables[13].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Respuesta:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, 
        "El principio SOLID que habla directamente en contra de esto es el Principio de Inversión de Dependencias (Dependency Inversion Principle - DIP), la letra 'D' de SOLID.\n\n"
        "¿Qué dice este principio?\n", bold=True
    )
    add_styled_run(p0, 
        "1. Los módulos de alto nivel (como TareaController, que contiene las reglas del negocio) no deben depender de módulos de bajo nivel (como ArrayList o detalles de almacenamiento). Ambos deben depender de abstracciones (interfaces como AlmacenarTareas).\n"
        "2. Las abstracciones no deben depender de los detalles concretos; los detalles deben depender de las abstracciones.\n\n"
        "Violación adicional:\n", bold=True
    )
    add_styled_run(p0, 
        "También vulnera el Principio de Responsabilidad Única (SRP), dado que el controlador asumía dos tareas distintas: orquestar la lógica de negocio y gestionar directamente la persistencia en memoria de las tareas."
    )

    # =========================================================================
    # Table 17: ¿Qué hace el @BeforeEach?
    # =========================================================================
    c = doc.tables[17].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Respuesta:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, 
        "@BeforeEach es una anotación fundamental de JUnit 5 (Jupiter) que indica que el método anotado debe ejecutarse obligatoriamente antes de cada método de prueba (@Test) individual dentro de la clase.\n\n"
        "Propósito y utilidad:\n", bold=True
    )
    add_styled_run(p0, 
        "• Configuración del Fixture: Permite instanciar objetos limpios y frescos (como un nuevo AlmacenamientoMemoria y un nuevo TareaController) para cada test.\n"
        "• Aislamiento e Independencia de Pruebas: Evita que el estado residual de un test (por ejemplo, tareas agregadas o modificadas en testAgregarVariosElementos) contamine o interfiera con la ejecución del siguiente test.\n"
        "• Reducción de código duplicado (DRY): Centraliza la inicialización de los componentes comunes sin necesidad de instanciarlos manualmente dentro de cada test."
    )

    # =========================================================================
    # Table 18: Tipos de assert que existen y ejemplos
    # =========================================================================
    c = doc.tables[18].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Tipos principales de assert en JUnit Jupiter (org.junit.jupiter.api.Assertions):\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, 
        "1. assertEquals(esperado, actual): Valida que ambos valores sean iguales.\n"
        "2. assertNotEquals(inesperado, actual): Valida que dos valores sean distintos.\n"
        "3. assertTrue(condicion) / assertFalse(condicion): Evalúa expresiones booleanas.\n"
        "4. assertNull(objeto) / assertNotNull(objeto): Verifica si una referencia es nula o no.\n"
        "5. assertSame(esperado, actual): Valida si ambas variables apuntan a la misma referencia en memoria (==).\n"
        "6. assertThrows(Excepcion.class, ejecutable): Valida que un bloque de código dispare la excepción requerida.\n"
        "7. assertAll(nombre, ejecutables...): Ejecuta un grupo de aserciones y reporta todos los fallos juntos.\n\n"
        "Ejemplo práctico incorporado a los tests:\n", bold=True
    )
    code_asserts = (
        "@Test\n"
        "void testDemostracionTiposDeAssert() {\n"
        "    Tarea t = controlador.agregarTarea(\"Practicar pruebas unitarias\", \"JUnit 5\");\n"
        "    \n"
        "    // assertNotNull: valida que el objeto fue instanciado\n"
        "    assertNotNull(t, \"La tarea no debe ser nula\");\n"
        "    \n"
        "    // assertEquals: valida el id asignado\n"
        "    assertEquals(1, t.getId());\n"
        "    \n"
        "    // assertFalse: valida que inicia sin completar\n"
        "    assertFalse(t.getCompletada());\n"
        "    \n"
        "    // assertAll: agrupa múltiples verificaciones garantizando que se evalúen todas\n"
        "    assertAll(\"Verificaciones del modelo Tarea\",\n"
        "        () -> assertEquals(1, t.getId()),\n"
        "        () -> assertFalse(t.getCompletada())\n"
        "    );\n"
        "}"
    )
    add_code_block(c, code_asserts)

    # =========================================================================
    # Table 19: Ventaja de recibir la interfaz en el constructor
    # =========================================================================
    c = doc.tables[19].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Respuesta:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, 
        "Las principales ventajas de recibir la interfaz AlmacenarTareas en el constructor (Inyección de Dependencias por Constructor) frente a instanciar new AlmacenamientoMemoria() internamente son:\n\n"
        "1. Desacoplamiento (Loose Coupling): ", bold=True
    )
    add_styled_run(p0, "El controlador desconoce cómo se almacenan las tareas; solo confía en el contrato de la interfaz. No está atado a una implementación rígida.\n")
    add_styled_run(p0, "2. Alta Testabilidad (Mocking): ", bold=True)
    add_styled_run(p0, "En pruebas unitarias complejas podemos inyectar un almacenamiento simulado (Mock o Stub) para probar el controlador de forma ultra rápida y sin efectos secundarios, sin requerir acceso a disco o bases de datos.\n")
    add_styled_run(p0, "3. Extensibilidad y Mantenimiento: ", bold=True)
    add_styled_run(p0, "Si el día de mañana se cambia a almacenamiento en base de datos PostgreSQL, Firebase o archivo JSON, no se toca una sola línea de TareaController; únicamente se le pasa la nueva instancia en su constructor al arrancar la aplicación.\n")
    add_styled_run(p0, "4. Cumplimiento de Principios SOLID: ", bold=True)
    add_styled_run(p0, "Respeta el Principio de Inversión de Dependencias (DIP) y el Principio Abierto/Cerrado (OCP).")

    # =========================================================================
    # Table 20: Cómo quedó instanciado TareaController en VistaTerminal / Main
    # =========================================================================
    c = doc.tables[20].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Instanciación en Main.java (Inyección de Dependencias en el punto de entrada):\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    code_main = (
        "package TAREAS;\n\n"
        "public class Main {\n"
        "    public static void main(String[] args) {\n"
        "        // Se crea la implementación concreta del almacenamiento\n"
        "        AlmacenarTareas almacenamiento = new AlmacenamientoMemoria();\n"
        "        \n"
        "        // Se inyecta la dependencia en el controlador\n"
        "        TareaController controlador = new TareaController(almacenamiento);\n"
        "        \n"
        "        // Se inyecta el controlador en la vista de terminal\n"
        "        VistaTerminal vista = new VistaTerminal(controlador);\n"
        "        vista.mostrarMenu();\n"
        "    }\n"
        "}"
    )
    add_code_block(c, code_main)
    
    p_vt = c.add_paragraph()
    style_paragraph(p_vt)
    add_styled_run(p_vt, "Y constructor correspondiente en VistaTerminal.java:\n", bold=True)
    code_vt_inst = (
        "public VistaTerminal() {\n"
        "    this.scanner = new Scanner(System.in);\n"
        "    AlmacenarTareas almacenamiento = new AlmacenamientoMemoria();\n"
        "    this.tareaController = new TareaController(almacenamiento);\n"
        "}\n\n"
        "public VistaTerminal(TareaController tareaController) {\n"
        "    this.scanner = new Scanner(System.in);\n"
        "    this.tareaController = tareaController;\n"
        "}"
    )
    add_code_block(c, code_vt_inst)

    # =========================================================================
    # Table 21: Screenshot de que los test corrieron (Fase 2)
    # =========================================================================
    c = doc.tables[21].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Captura de pantalla: Pruebas unitarias de la fase 2 tras el desacoplamiento con AlmacenarTareas y @BeforeEach:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    p_img = c.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_paragraph(p_img, space_before=4, space_after=4)
    p_img.add_run().add_picture('assets/screenshot_part2.png', width=Inches(5.7))

    # =========================================================================
    # Table 22: Método boolean eliminar(int id) en interfaz y en AlmacenamientoMemoria
    # =========================================================================
    c = doc.tables[22].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "1. En la interfaz AlmacenarTareas.java:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    code_elim_int = (
        "public interface AlmacenarTareas {\n"
        "    void guardar(Tarea tarea);\n"
        "    Tarea buscarPorId(int id);\n"
        "    List<Tarea> obtenerTodas();\n"
        "    boolean eliminar(int id); // Nuevo método agregado al contrato\n"
        "}"
    )
    add_code_block(c, code_elim_int)

    p1 = c.add_paragraph()
    style_paragraph(p1)
    add_styled_run(p1, "2. Implementación en AlmacenamientoMemoria.java:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    code_elim_mem = (
        "@Override\n"
        "public boolean eliminar(int id) {\n"
        "    Tarea t = buscarPorId(id);\n"
        "    if (t != null) {\n"
        "        return listaDeTareas.remove(t);\n"
        "    }\n"
        "    return false;\n"
        "}"
    )
    add_code_block(c, code_elim_mem)

    # =========================================================================
    # Table 23: Método public void eliminarTarea(int id) en TareaController
    # =========================================================================
    c = doc.tables[23].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Método eliminarTarea en TareaController.java con validación de excepción:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    code_elim_ctrl = (
        "public void eliminarTarea(int id) {\n"
        "    Tarea t = almacenamiento.buscarPorId(id);\n"
        "    if (t == null) {\n"
        "        throw new IllegalArgumentException(\"No hay una tarea con el ID: \" + id);\n"
        "    }\n"
        "    almacenamiento.eliminar(id);\n"
        "}"
    )
    add_code_block(c, code_elim_ctrl)

    # =========================================================================
    # Table 24: Dos tests para eliminar tarea en TareaControllerTest.java
    # =========================================================================
    c = doc.tables[24].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Tests unitarios en TareaControllerTest.java para eliminar tareas:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    code_test_elim = (
        "@Test\n"
        "void testEliminarTareaExistente() {\n"
        "    Tarea t = controlador.agregarTarea(\"Comprar insumos\", \"Papeleria\");\n"
        "    int id = t.getId();\n"
        "    \n"
        "    controlador.eliminarTarea(id); // Se elimina la tarea creada\n"
        "    \n"
        "    // Confirmamos que ya no existe verificando que intentar completarla lance la excepción\n"
        "    assertThrows(IllegalArgumentException.class, () -> {\n"
        "        controlador.completarTarea(id);\n"
        "    });\n"
        "}\n\n"
        "@Test\n"
        "void testEliminarTareaInexistenteLanzaExcepcion() {\n"
        "    // Intentar eliminar un ID no existente debe lanzar IllegalArgumentException\n"
        "    assertThrows(IllegalArgumentException.class, () -> {\n"
        "        controlador.eliminarTarea(999);\n"
        "    });\n"
        "}"
    )
    add_code_block(c, code_test_elim)

    # =========================================================================
    # Table 25: Opción "Eliminar tarea" en VistaTerminal.java con try-catch
    # =========================================================================
    c = doc.tables[25].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Implementación de eliminarTarea() en VistaTerminal.java con manejo de errores:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    code_elim_vt = (
        "private void eliminarTarea() {\n"
        "    System.out.println(\"\\n--- Eliminar Tarea ---\");\n"
        "    System.out.print(\"Ingrese el ID de la tarea a eliminar: \");\n"
        "    try {\n"
        "        int id = scanner.nextInt();\n"
        "        scanner.nextLine(); // Limpiar el buffer\n"
        "        tareaController.eliminarTarea(id);\n"
        "        System.out.println(\"Tarea con ID \" + id + \" eliminada exitosamente.\");\n"
        "    } catch (IllegalArgumentException e) {\n"
        "        System.out.println(\"Error: \" + e.getMessage());\n"
        "    } catch (Exception e) {\n"
        "        System.out.println(\"Entrada inválida. Debe ingresar un número entero.\");\n"
        "        scanner.nextLine(); // Descartar entrada incorrecta\n"
        "    }\n"
        "}"
    )
    add_code_block(c, code_elim_vt)

    # =========================================================================
    # Table 26: Código completo de la funcionalidad y tests JUnit
    # =========================================================================
    c = doc.tables[26].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Código completo de la suite de pruebas TareaControllerTest.java:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    code_full_tests = (
        "package TAREAS;\n\n"
        "import static org.junit.jupiter.api.Assertions.*;\n"
        "import org.junit.jupiter.api.BeforeEach;\n"
        "import org.junit.jupiter.api.Test;\n\n"
        "public class TareaControllerTest {\n"
        "    private TareaController controlador;\n\n"
        "    @BeforeEach\n"
        "    void setUp() {\n"
        "        AlmacenarTareas almacenamiento = new AlmacenamientoMemoria();\n"
        "        controlador = new TareaController(almacenamiento);\n"
        "    }\n\n"
        "    @Test\n"
        "    void testAgregarVariosElementosYRevisarIds() {\n"
        "        Tarea t1 = controlador.agregarTarea(\"atender estudiantes\", \"Ximena\");\n"
        "        Tarea t2 = controlador.agregarTarea(\"atender estudiantes\", \"Luis Diego\");\n"
        "        Tarea t3 = controlador.agregarTarea(\"atender estudiantes\", \"william\");\n"
        "        assertEquals(1, t1.getId());\n"
        "        assertEquals(2, t2.getId());\n"
        "        assertEquals(3, t3.getId());\n"
        "    }\n\n"
        "    @Test\n"
        "    void validarCambioDeEstado() {\n"
        "        Tarea t = controlador.agregarTarea(\"atender estudiantes\", \"william\");\n"
        "        assertFalse(t.getCompletada());\n"
        "        controlador.completarTarea(t.getId());\n"
        "        assertTrue(t.getCompletada());\n"
        "    }\n\n"
        "    @Test\n"
        "    void testCompletarTareaInexistenteLanzaExcepcion() {\n"
        "        assertThrows(IllegalArgumentException.class, () -> {\n"
        "            controlador.completarTarea(99);\n"
        "        });\n"
        "    }\n\n"
        "    @Test\n"
        "    void testCompletarTareaYaCompletadaLanzaExcepcion() {\n"
        "        Tarea t = controlador.agregarTarea(\"Estudiar para el examen\", \"POO\");\n"
        "        controlador.completarTarea(t.getId());\n"
        "        assertThrows(IllegalStateException.class, () -> {\n"
        "            controlador.completarTarea(t.getId());\n"
        "        });\n"
        "    }\n\n"
        "    @Test\n"
        "    void testEliminarTareaExistente() {\n"
        "        Tarea t = controlador.agregarTarea(\"Comprar insumos\", \"Papeleria\");\n"
        "        int id = t.getId();\n"
        "        controlador.eliminarTarea(id);\n"
        "        assertThrows(IllegalArgumentException.class, () -> {\n"
        "            controlador.completarTarea(id);\n"
        "        });\n"
        "    }\n\n"
        "    @Test\n"
        "    void testEliminarTareaInexistenteLanzaExcepcion() {\n"
        "        assertThrows(IllegalArgumentException.class, () -> {\n"
        "            controlador.eliminarTarea(999);\n"
        "        });\n"
        "    }\n"
        "}"
    )
    add_code_block(c, code_full_tests)

    # =========================================================================
    # Table 27: Screenshot de que los test corrieron (Fase 3 - Final Completa)
    # =========================================================================
    c = doc.tables[27].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Captura de pantalla: Todas las pruebas unitarias (100% pasando en verde):\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    p_img = c.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_paragraph(p_img, space_before=4, space_after=4)
    p_img.add_run().add_picture('assets/screenshot_part3.png', width=Inches(5.7))

    # =========================================================================
    # Table 28: Pregunta final: TDD y cómo se relaciona con lo realizado
    # =========================================================================
    c = doc.tables[28].rows[0].cells[0]
    clear_cell(c)
    p0 = c.paragraphs[0]
    style_paragraph(p0)
    add_styled_run(p0, "Respuesta:\n", bold=True, color=RGBColor(0x1F, 0x49, 0x7D))
    add_styled_run(p0, 
        "¿Qué es Test-Driven Development (TDD)?\n", bold=True
    )
    add_styled_run(p0, 
        "TDD (Desarrollo Guiado por Pruebas) es una metodología de ingeniería de software basada en iteraciones muy cortas donde las pruebas unitarias automatizadas se diseñan y escriben ANTES del código de producción.\n\n"
        "Sigue el ciclo fundamental Red - Green - Refactor:\n"
        "1. Red (Rojo): Se escribe una prueba unitaria para definir una nueva funcionalidad requerida. Al ejecutarse, la prueba falla porque la funcionalidad aún no existe.\n"
        "2. Green (Verde): Se escribe el código de producción mínimo y necesario para hacer que la prueba pase con éxito.\n"
        "3. Refactor (Refactorizar): Se reestructura, limpia y optimiza el diseño del código (eliminando redundancias y aplicando principios de diseño como SOLID), teniendo la garantía de que los tests siguen pasando en verde.\n\n"
        "¿Cómo se relaciona con lo que acabamos de realizar en el laboratorio?\n", bold=True
    )
    add_styled_run(p0, 
        "1. Diseño guiado por especificaciones: En el lab, antes de implementar la interfaz gráfica o de terminal final, primero definimos y aseguramos las reglas de negocio a través de tests en TareaControllerTest (por ejemplo: completarTarea, verificar cambio de boolean, validar excepciones con assertThrows ante tareas inexistentes o ya completadas).\n"
        "2. Red de Seguridad en el Refactor: Cuando desacoplamos TareaController para que no dependiera de ArrayList sino de la interfaz AlmacenarTareas, los tests ya creados actuaron como una red de protección inmediata (safety net). Ejecutar las pruebas nos garantizó al instante que la reestructuración de la arquitectura no rompió la lógica existente.\n"
        "3. Incorporación de nueva funcionalidad (Eliminar Tarea): Al agregar la eliminación, definimos los casos de prueba (testEliminarTareaExistente y testEliminarTareaInexistenteLanzaExcepcion) para especificar con exactitud el comportamiento deseado de los métodos antes de conectarlos al menú de la terminal."
    )

    doc.save('lab iv.docx')
    print("Document successfully updated and saved as 'lab iv.docx'!")

if __name__ == '__main__':
    fill_doc()

