package TAREAS;

import java.util.List;
import java.util.Scanner;

public class VistaTerminal {
    private Scanner scanner;
    private TareaController tareaController;

    public VistaTerminal() {
        this.scanner = new Scanner(System.in);
        AlmacenarTareas almacenamiento = new AlmacenamientoMemoria();
        this.tareaController = new TareaController(almacenamiento);
    }

    public VistaTerminal(TareaController tareaController) {
        this.scanner = new Scanner(System.in);
        this.tareaController = tareaController;
    }

    public void mostrarMenu() {
        boolean salir = false;
        while (!salir) {
            System.out.println("\n========================================");
            System.out.println("Bienvenido al sistema de gestión de tareas");
            System.out.println("========================================");
            System.out.println("Seleccione una opción:");
            System.out.println("1. Agregar tarea");
            System.out.println("2. Completar tarea");
            System.out.println("3. Eliminar tarea");
            System.out.println("4. Listar tareas");
            System.out.println("5. Salir");
            System.out.print("Opción: ");

            int opcion = -1;
            try {
                opcion = scanner.nextInt();
                scanner.nextLine(); // Limpiar el buffer
            } catch (Exception e) {
                System.out.println("Entrada inválida. Debe ingresar un número entero.");
                scanner.nextLine();
                continue;
            }

            switch (opcion) {
                case 1:
                    agregarTarea();
                    break;
                case 2:
                    completarTarea();
                    break;
                case 3:
                    eliminarTarea();
                    break;
                case 4:
                    listarTareas();
                    break;
                case 5:
                    salir = true;
                    System.out.println("¡Hasta luego!");
                    break;
                default:
                    System.out.println("Opción inválida. Intente nuevamente.");
            }
        }
    }

    private void agregarTarea() {
        System.out.println("\n--- Agregar Tarea ---");
        System.out.print("Ingrese el título de la tarea: ");
        String titulo = scanner.nextLine();
        System.out.print("Ingrese la descripción: ");
        String descripcion = scanner.nextLine();

        Tarea nueva = tareaController.agregarTarea(titulo, descripcion);
        System.out.println("Tarea creada exitosamente con ID: " + nueva.getId());
    }

    private void completarTarea() {
        System.out.println("\n--- Completar Tarea ---");
        System.out.print("Ingrese el ID de la tarea a completar: ");
        try {
            int id = scanner.nextInt();
            scanner.nextLine();
            tareaController.completarTarea(id);
            System.out.println("Tarea con ID " + id + " completada exitosamente.");
        } catch (IllegalArgumentException | IllegalStateException e) {
            System.out.println("Error: " + e.getMessage());
        } catch (Exception e) {
            System.out.println("Entrada inválida. Debe ingresar un número entero.");
            scanner.nextLine();
        }
    }

    private void eliminarTarea() {
        System.out.println("\n--- Eliminar Tarea ---");
        System.out.print("Ingrese el ID de la tarea a eliminar: ");
        try {
            int id = scanner.nextInt();
            scanner.nextLine();
            tareaController.eliminarTarea(id);
            System.out.println("Tarea con ID " + id + " eliminada exitosamente.");
        } catch (IllegalArgumentException e) {
            System.out.println("Error: " + e.getMessage());
        } catch (Exception e) {
            System.out.println("Entrada inválida. Debe ingresar un número entero.");
            scanner.nextLine();
        }
    }

    private void listarTareas() {
        System.out.println("\n--- Lista de Tareas ---");
        List<Tarea> tareas = tareaController.getListaDeTareas();
        if (tareas.isEmpty()) {
            System.out.println("No hay tareas registradas.");
        } else {
            for (Tarea t : tareas) {
                String estado = t.getCompletada() ? "[COMPLETADA]" : "[PENDIENTE]";
                System.out.println("ID: " + t.getId() + " " + estado);
            }
        }
    }
}
