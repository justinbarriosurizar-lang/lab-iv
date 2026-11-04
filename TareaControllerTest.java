package TAREAS;

import static org.junit.jupiter.api.Assertions.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

public class TareaControllerTest {
    
    private TareaController controlador;

    @BeforeEach
    void setUp() {
        AlmacenarTareas almacenamiento = new AlmacenamientoMemoria();
        controlador = new TareaController(almacenamiento);
    }

    @Test
    void testAgregarVariosElementosYRevisarIds() {
        Tarea t1 = controlador.agregarTarea("atender estudiantes", "Ximena");
        Tarea t2 = controlador.agregarTarea("atender estudiantes", "Luis Diego");
        Tarea t3 = controlador.agregarTarea("atender estudiantes", "william");
        
        assertEquals(1, t1.getId());
        assertEquals(2, t2.getId());
        assertEquals(3, t3.getId());
    }

    @Test
    void validarCambioDeEstado() {
        Tarea t = controlador.agregarTarea("atender estudiantes", "william");
        assertFalse(t.getCompletada());
        
        controlador.completarTarea(t.getId());
        assertTrue(t.getCompletada());
    }

    @Test
    void testCompletarTareaInexistenteLanzaExcepcion() {
        assertThrows(IllegalArgumentException.class, () -> {
            controlador.completarTarea(99);
        });
    }

    @Test
    void testCompletarTareaYaCompletadaLanzaExcepcion() {
        Tarea t = controlador.agregarTarea("Estudiar para el examen", "POO");
        controlador.completarTarea(t.getId());
        
        assertThrows(IllegalStateException.class, () -> {
            controlador.completarTarea(t.getId());
        });
    }

    @Test
    void testEliminarTareaExistente() {
        Tarea t = controlador.agregarTarea("Comprar insumos", "Papeleria");
        int id = t.getId();
        
        controlador.eliminarTarea(id);
        
        assertThrows(IllegalArgumentException.class, () -> {
            controlador.completarTarea(id);
        });
    }

    @Test
    void testEliminarTareaInexistenteLanzaExcepcion() {
        assertThrows(IllegalArgumentException.class, () -> {
            controlador.eliminarTarea(999);
        });
    }

    @Test
    void testDemostracionTiposDeAssert() {
        Tarea t = controlador.agregarTarea("Practicar pruebas unitarias", "JUnit 5");
        
        // 1. assertNotNull
        assertNotNull(t, "La tarea no debería ser nula");
        
        // 2. assertEquals
        assertEquals("Practicar pruebas unitarias", "Practicar pruebas unitarias");
        
        // 3. assertFalse
        assertFalse(t.getCompletada());
        
        // 4. assertAll (evalúa múltiples aserciones en conjunto)
        assertAll("Verificación completa de la tarea",
            () -> assertEquals(1, t.getId()),
            () -> assertFalse(t.getCompletada())
        );
    }
}
