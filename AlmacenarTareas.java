package TAREAS;

import java.util.List;

public interface AlmacenarTareas {
    void guardar(Tarea tarea);
    Tarea buscarPorId(int id);
    List<Tarea> obtenerTodas();
    boolean eliminar(int id);
}

