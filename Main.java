package TAREAS;

public class Main {
    public static void main(String[] args) {
        AlmacenarTareas almacenamiento = new AlmacenamientoMemoria();
        TareaController controlador = new TareaController(almacenamiento);
        VistaTerminal vista = new VistaTerminal(controlador);
        vista.mostrarMenu();
    }
}
