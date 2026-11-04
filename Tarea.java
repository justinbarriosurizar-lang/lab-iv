package TAREAS;

public class Tarea {
    private int id;
    private String titulo;
    private String descripcion;
    private boolean completada;

    public Tarea(int id, String t, String c) {
        this.id = id;
        this.titulo = t;
        this.descripcion = c;
        this.completada = false;
    }

    public int getId() {
        return this.id;
    }

    public void setCompletada(boolean estadoNuevo) {
        this.completada = estadoNuevo;
    }

    public boolean getCompletada() {
        return this.completada;
    }
}