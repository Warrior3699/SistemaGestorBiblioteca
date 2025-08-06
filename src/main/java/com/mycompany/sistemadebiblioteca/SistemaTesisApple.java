/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */
package com.mycompany.sistemadebiblioteca;

import javax.swing.*;
import javax.swing.border.*;
import javax.swing.table.*;
import java.awt.*;
import java.awt.event.*;
import java.sql.*;
import java.util.*;
import java.io.File;
import java.io.FileOutputStream;

// Apache POI imports
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.FillPatternType;

public class SistemaTesisApple {

    private Connection conn;
    private JFrame currentFrame;
    private String currentUser;
    private DefaultTableModel tableModel;
    private JTable tabla;
    private final Color COLOR_PRIMARIO = new Color(0, 122, 255);
    private final Color COLOR_ELIMINAR = new Color(255, 59, 48);
    private final Color FONDO = new Color(242, 242, 247);
    private final Color CARD_BG = Color.WHITE;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new SistemaTesisApple().iniciar());
    }

    private void iniciar() {
        conectarBD();
        mostrarAuth();
    }

    private void conectarBD() {
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            conn = DriverManager.getConnection(
                    "jdbc:mysql://localhost:3306/gestion_tesis_apple",
                    "root",
                    "1234"
            );
            crearTablas();
        } catch (Exception e) {
            mostrarError("Error de conexión: " + e.getMessage());
            System.exit(1);
        }
    }

    private void crearTablas() {
        String[] queries = {
            "CREATE TABLE IF NOT EXISTS usuarios ("
            + "  id INT AUTO_INCREMENT PRIMARY KEY,"
            + "  nombre VARCHAR(50) NOT NULL,"
            + "  usuario VARCHAR(50) UNIQUE NOT NULL,"
            + "  contrasena VARCHAR(100) NOT NULL"
            + ")",
            "CREATE TABLE IF NOT EXISTS tesis ("
            + "  id INT AUTO_INCREMENT PRIMARY KEY,"
            + "  numero INT NULL,"
            + "  titulo VARCHAR(255) NOT NULL,"
            + "  licenciatura VARCHAR(100) NOT NULL,"
            + "  autor VARCHAR(100) NOT NULL,"
            + "  asesor VARCHAR(100) NOT NULL,"
            + "  estado VARCHAR(50) DEFAULT 'En progreso',"
            + "  mes VARCHAR(20)  NULL,"
            + "  anio INT         NULL,"
            + "  usuario_id INT,"
            + "  FOREIGN KEY (usuario_id) REFERENCES usuarios(id)"
            + ")"
        };

        try (Statement stmt = conn.createStatement()) {
            for (String sql : queries) stmt.execute(sql);
        } catch (SQLException e) {
            System.err.println("Error creando tablas: " + e.getMessage());
        }
    }

    // ====== AUTENTICACIÓN ======
    private void mostrarAuth() {
        currentFrame = new JFrame("Tesla Tesis");
        currentFrame.setSize(500, 600);
        currentFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        currentFrame.setLayout(new BorderLayout());
        currentFrame.getContentPane().setBackground(FONDO);

        JPanel header = crearPanelHeader("Tesla Tesis", "Repositorio Académico Colaborativo");
        currentFrame.add(header, BorderLayout.NORTH);

        JPanel card = new JPanel(new GridLayout(0, 1, 10, 10));
        card.setBackground(CARD_BG);
        card.setBorder(BorderFactory.createEmptyBorder(30, 40, 30, 40));

        JTextField txtUsuario = crearCampoTexto("Usuario");
        JPasswordField txtContrasena = crearCampoTextoContrasena();

        card.add(crearEtiqueta("Usuario"));
        card.add(txtUsuario);
        card.add(crearEtiqueta("Contraseña"));
        card.add(txtContrasena);

        JButton btnLogin = crearBotonPrimario("Iniciar Sesión", e
                -> validarUsuario(txtUsuario.getText().trim(),
                        new String(txtContrasena.getPassword()))
        );
        JButton btnRegistro = crearBotonSecundario("Crear nueva cuenta", e
                -> mostrarRegistro()
        );

        card.add(btnLogin);
        card.add(btnRegistro);

        currentFrame.add(card, BorderLayout.CENTER);
        centrarVentana(currentFrame);
        currentFrame.setVisible(true);
    }

    private void validarUsuario(String usuario, String contrasena) {
        if (usuario.isEmpty() || contrasena.isEmpty()) {
            mostrarError("Usuario y contraseña requeridos");
            return;
        }

        String sql = "SELECT id, nombre, usuario FROM usuarios WHERE usuario = ? AND contrasena = ?";
        try (PreparedStatement stmt = conn.prepareStatement(sql)) {
            stmt.setString(1, usuario);
            stmt.setString(2, contrasena);
            ResultSet rs = stmt.executeQuery();
            if (rs.next()) {
                currentUser = rs.getString("usuario");
                mostrarDashboard();
            } else {
                mostrarError("Credenciales incorrectas");
            }
        } catch (SQLException e) {
            mostrarError("Error de autenticación: " + e.getMessage());
        }
    }

    private void registrarUsuario(String nombre, String usuario, String contrasena) {
        if (nombre.isEmpty() || usuario.isEmpty() || contrasena.isEmpty()) {
            mostrarError("Todos los campos son obligatorios");
            return;
        }
        if (contrasena.length() < 6) {
            mostrarError("La contraseña debe tener al menos 6 caracteres");
            return;
        }

        String sql = "INSERT INTO usuarios (nombre, usuario, contrasena) VALUES (?, ?, ?)";
        try (PreparedStatement stmt = conn.prepareStatement(sql)) {
            stmt.setString(1, nombre);
            stmt.setString(2, usuario);
            stmt.setString(3, contrasena);
            if (stmt.executeUpdate() > 0) {
                mostrarMensaje("¡Cuenta creada con éxito!");
                currentUser = usuario;
                mostrarDashboard();
            }
        } catch (SQLException e) {
            if (e.getMessage().contains("Duplicate")) {
                mostrarError("El nombre de usuario ya existe");
            } else {
                mostrarError("Error al registrar: " + e.getMessage());
            }
        }
    }

    private void mostrarRegistro() {
        JDialog dialog = new JDialog(currentFrame, "Nueva Cuenta", true);
        dialog.setLayout(new BorderLayout());
        dialog.getContentPane().setBackground(FONDO);

        JPanel header = crearPanelHeader("Nueva Cuenta", "Únete a la comunidad académica");
        dialog.add(header, BorderLayout.NORTH);

        JPanel card = new JPanel(new GridLayout(0, 2, 10, 10));
        card.setBackground(CARD_BG);
        card.setBorder(BorderFactory.createEmptyBorder(30, 40, 30, 40));

        JTextField txtNombre = crearCampoTexto("Nombre completo");
        JTextField txtUsuario = crearCampoTexto("Usuario");
        JPasswordField txtContrasena = crearCampoTextoContrasena();

        card.add(crearEtiqueta("Nombre"));
        card.add(txtNombre);
        card.add(crearEtiqueta("Usuario"));
        card.add(txtUsuario);
        card.add(crearEtiqueta("Contraseña"));
        card.add(txtContrasena);

        JButton btnRegistrar = crearBotonPrimario("Crear Cuenta", e -> {
            registrarUsuario(txtNombre.getText().trim(),
                    txtUsuario.getText().trim(),
                    new String(txtContrasena.getPassword()));
            dialog.dispose();
        });

        card.add(new JLabel());
        card.add(btnRegistrar);

        dialog.add(card, BorderLayout.CENTER);
        centrarVentana(dialog);
        dialog.pack();
        dialog.setVisible(true);
    }

    // ====== DASHBOARD ======
    private void mostrarDashboard() {
        currentFrame.dispose();
        currentFrame = new JFrame("Tesla Tesis - " + currentUser);
        currentFrame.setSize(1200, 800);
        currentFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        currentFrame.setLayout(new BorderLayout());
        currentFrame.getContentPane().setBackground(FONDO);

        JPanel topBar = new JPanel(new BorderLayout());
        topBar.setBackground(CARD_BG);
        topBar.setBorder(BorderFactory.createEmptyBorder(15, 25, 15, 25));
        JLabel lblTitulo = new JLabel("Repositorio de Tesis");
        lblTitulo.setFont(new Font("SF Pro Display", Font.BOLD, 24));
        topBar.add(lblTitulo, BorderLayout.WEST);

        JPanel topBtns = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));
        topBtns.setBackground(CARD_BG);
        JButton btnNueva = crearBotonPrimario("+ Nueva Tesis", e -> mostrarFormularioTesis());
        JButton btnExportar = crearBotonSecundario("Exportar a Excel", e -> exportarExcel());
        JButton btnLogout = crearBotonSecundario("Cerrar Sesión", e -> {
            currentUser = null;
            mostrarAuth();
        });
        topBtns.add(btnNueva);
        topBtns.add(btnExportar);
        topBtns.add(btnLogout);
        topBar.add(topBtns, BorderLayout.EAST);

        JPanel searchPanel = new JPanel(new BorderLayout());
        searchPanel.setBackground(CARD_BG);
        searchPanel.setBorder(BorderFactory.createEmptyBorder(15, 25, 15, 25));

        JTextField txtBusqueda = crearCampoTexto("Buscar por título, autor, asesor o licenciatura...");
        txtBusqueda.addActionListener(e -> buscarTesis(txtBusqueda.getText().trim()));

        JButton btnBuscar = crearBotonIcono("🔍");
        btnBuscar.addActionListener(e -> buscarTesis(txtBusqueda.getText().trim()));

        JButton btnBusqAvanzada = crearBotonIcono("Busqueda Avanzada");
        btnBusqAvanzada.addActionListener(e -> mostrarBusquedaAvanzada());

        JPanel grp = new JPanel(new BorderLayout(10, 0));
        grp.add(txtBusqueda, BorderLayout.CENTER);
        grp.add(btnBuscar, BorderLayout.EAST);
        grp.add(btnBusqAvanzada, BorderLayout.SOUTH);
        searchPanel.add(grp, BorderLayout.CENTER);

        configurarTabla();
        cargarTesis("");

        JScrollPane scroll = new JScrollPane(tabla);
        scroll.setBorder(BorderFactory.createEmptyBorder());

        currentFrame.add(topBar, BorderLayout.NORTH);
        currentFrame.add(searchPanel, BorderLayout.CENTER);
        currentFrame.add(scroll, BorderLayout.SOUTH);
        centrarVentana(currentFrame);
        currentFrame.setVisible(true);
    }

    private void configurarTabla() {
        String[] columnas = {
            "ID", "Número", "Título", "Licenciatura", "Autor", "Asesor",
            "Estado", "Mes", "Año", "Subido por", "Acciones"
        };
        tableModel = new DefaultTableModel(columnas, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return column == 10;
            }
        };
        tabla = new JTable(tableModel);
        estilizarTabla(tabla);

        tabla.getColumnModel().getColumn(10)
                .setCellRenderer(new ButtonRenderer(currentUser));
        tabla.getColumnModel().getColumn(10)
                .setCellEditor(new ButtonEditor(new JCheckBox(), currentUser));
    }

    private void buscarEnBaseDeDatos(String numero, String titulo, String licenciatura, String autor, String asesor, String mes, String anio) {
        tableModel.setRowCount(0);
        StringBuilder query = new StringBuilder("SELECT t.id, t.numero, t.titulo, t.licenciatura, t.autor, t.asesor, t.estado, t.mes, t.anio, u.usuario AS subido_por FROM tesis t "
            + "JOIN usuarios u ON t.usuario_id = u.id WHERE 1=1");
        if (!numero.isEmpty()) {
            query.append(" AND t.numero = " + numero);
        }
        if (!titulo.isEmpty()) {
            query.append(" AND t.titulo = '" + titulo + "'");
        }
        if (!licenciatura.isEmpty()) {
            query.append(" AND t.licenciatura = '" + licenciatura + "'");
        }
        if (!autor.isEmpty()) {
            query.append(" AND t.autor = '" + autor + "'");
        }
        if (!asesor.isEmpty()) {
            query.append(" AND t.asesor = '" + asesor + "'");
        }
        if (!mes.isEmpty()) {
            query.append(" AND t.mes = '" + mes + "'");
        }
        if (!anio.isEmpty()) {
            query.append(" AND t.anio = " + anio);
        }
        query.append(" ORDER BY COALESCE(t.anio,0) DESC, t.numero ASC");

        String sql = query.toString();
        try (PreparedStatement stmt = conn.prepareStatement(sql)) {
            ResultSet rs = stmt.executeQuery();
            while (rs.next()) {
                String numeroStr = rs.getString("numero");
                String tituloStr = rs.getString("titulo");
                String licenciaturaStr = rs.getString("licenciatura");
                String autorStr = rs.getString("autor");
                String asesorStr = rs.getString("asesor");
                String estado = rs.getString("estado");
                String mesStr = rs.getString("mes");
                String anioStr = rs.getString("anio");
                String subidoPor = rs.getString("subido_por");

                tableModel.addRow(new Object[]{
                    rs.getInt("id"),
                    numeroStr,
                    tituloStr,
                    licenciaturaStr,
                    autorStr,
                    asesorStr,
                    estado,
                    mesStr,
                    anioStr == null ? "" : anioStr,
                    subidoPor,
                    "✎ Editar | 🗑 Eliminar"
                });
            }
        } catch (SQLException e) {
            mostrarError("Error al cargar tesis: " + e.getMessage());
        }
    }

    private void cargarTesis(String busqueda) {
        tableModel.setRowCount(0);
        String sql
                = "SELECT t.id, t.numero, t.titulo, t.licenciatura, t.autor, t.asesor, "
                + "       t.estado, t.mes, t.anio, u.usuario AS subido_por "
                + "FROM tesis t "
                + "JOIN usuarios u ON t.usuario_id = u.id "
                + "WHERE t.titulo LIKE ? OR t.autor LIKE ? OR t.asesor LIKE ? OR t.licenciatura LIKE ? "
                + "ORDER BY COALESCE(t.anio,0) DESC, t.numero ASC";

        try (PreparedStatement stmt = conn.prepareStatement(sql)) {
            String p = "%" + busqueda + "%";
            stmt.setString(1, p);
            stmt.setString(2, p);
            stmt.setString(3, p);
            stmt.setString(4, p);

            ResultSet rs = stmt.executeQuery();
            while (rs.next()) {
                String numero = rs.getString("numero");
                String titulo = rs.getString("titulo");
                String licenciatura = rs.getString("licenciatura");
                String autor = rs.getString("autor");
                String asesor = rs.getString("asesor");
                String estado = rs.getString("estado");
                String mes = rs.getString("mes");
                Object anioObj = rs.getObject("anio");  // puede ser null
                String subidoPor = rs.getString("subido_por");

                tableModel.addRow(new Object[]{
                    rs.getInt("id"),
                    numero,
                    titulo,
                    licenciatura,
                    autor,
                    asesor,
                    estado,
                    mes,
                    anioObj == null ? "" : anioObj.toString(),
                    subidoPor,
                    "✎ Editar | 🗑 Eliminar"
                });
            }
        } catch (SQLException e) {
            mostrarError("Error al cargar tesis: " + e.getMessage());
        }
    }

    private void buscarTesis(String busqueda) {
        cargarTesis(busqueda);
    }

    // ====== FORMULARIO CRUD ======
    private void mostrarFormularioTesis() {
        mostrarFormularioTesis(-1, null);
    }

    private void mostrarBusquedaAvanzada() {
        JDialog dialoga = new JDialog(currentFrame);
        dialoga.setSize(900, 800);
        dialoga.setLayout(new BorderLayout());
        dialoga.getContentPane().setBackground(FONDO);

        JPanel header = new JPanel(new BorderLayout());
        header.setBackground(FONDO);
        header.setBorder(BorderFactory.createEmptyBorder(20, 20, 10, 20));
        JLabel lbl = new JLabel("Busqueda avanzada");
        lbl.setFont(new Font("SF Pro Display", Font.BOLD, 24));
        header.add(lbl, BorderLayout.CENTER);

        JPanel card = new JPanel(new GridLayout(0, 1, 10, 15));
        card.setBackground(CARD_BG);
        card.setBorder(BorderFactory.createEmptyBorder(10, 20, 10, 20));

        JTextField txtNumero = crearCampoTexto("Número de la tesis");
        JTextField txtTitulo = crearCampoTexto("Título");
        JTextField txtLicenciatura = crearCampoTexto("Licenciatura");
        JTextField txtAutor = crearCampoTexto("Autor");
        JTextField txtAsesor = crearCampoTexto("Asesor");
        JTextField txtMes = crearCampoTexto("Mes");
        JTextField txtAnio = crearCampoTexto("Año");

        card.add(crearEtiqueta("Número"));
        card.add(txtNumero);
        card.add(crearEtiqueta("Título"));
        card.add(txtTitulo);
        card.add(crearEtiqueta("Licenciatura"));
        card.add(txtLicenciatura);
        card.add(crearEtiqueta("Autor"));
        card.add(txtAutor);
        card.add(crearEtiqueta("Asesor"));
        card.add(txtAsesor);
        card.add(crearEtiqueta("Mes"));
        card.add(txtMes);
        card.add(crearEtiqueta("Año"));
        card.add(txtAnio);

        card.add(Box.createVerticalStrut(20));
        JButton btnBusquedapro = crearBotonPrimario("Buscar", null);
        card.add(btnBusquedapro);
        btnBusquedapro.addActionListener(e -> {
            buscarEnBaseDeDatos(
                    txtNumero.getText().trim(),
                    txtTitulo.getText().trim(),
                    txtLicenciatura.getText().trim(),
                    txtAutor.getText().trim(),
                    txtAsesor.getText().trim(),
                    txtMes.getText().trim(),
                    txtAnio.getText().trim()
            );
        });

        dialoga.add(header, BorderLayout.NORTH);
        dialoga.add(new JScrollPane(card), BorderLayout.CENTER);
        centrarVentana(dialoga);
        dialoga.setVisible(true);
    }

    private void mostrarFormularioTesis(int idTesis, String usuarioTesis) {
        JDialog dialog = new JDialog(currentFrame,
                idTesis == -1 ? "Nueva Tesis" : "Editar Tesis", true);
        dialog.setSize(600, 800);
        dialog.setLayout(new BorderLayout());
        dialog.getContentPane().setBackground(FONDO);

        JPanel header = new JPanel(new BorderLayout());
        header.setBackground(FONDO);
        header.setBorder(BorderFactory.createEmptyBorder(20, 20, 10, 20));
        JLabel lbl = new JLabel(idTesis == -1 ? "Nueva Tesis" : "Editar Tesis");
        lbl.setFont(new Font("SF Pro Display", Font.BOLD, 24));
        header.add(lbl, BorderLayout.CENTER);
        if (idTesis != -1) {
            JLabel sub = new JLabel("ID: " + idTesis + " | Subido por: " + usuarioTesis);
            sub.setFont(new Font("SF Pro Text", Font.PLAIN, 12));
            sub.setForeground(Color.GRAY);
            header.add(sub, BorderLayout.SOUTH);
        }

        JPanel card = new JPanel(new GridLayout(0, 1, 10, 15));
        card.setBackground(CARD_BG);
        card.setBorder(BorderFactory.createEmptyBorder(30, 40, 30, 40));

        JTextField txtNumero = crearCampoTexto("Número de la tesis");
        JTextField txtTitulo = crearCampoTexto("Título");
        JTextField txtLicenciatura = crearCampoTexto("Licenciatura");
        JTextField txtAutor = crearCampoTexto("Autor");
        JTextField txtAsesor = crearCampoTexto("Asesor");
        JTextField txtMes = crearCampoTexto("Mes");
        JTextField txtAnio = crearCampoTexto("Año");

        JComboBox<String> cmbEstado = new JComboBox<>(
                new String[]{"En progreso", "Finalizada", "En revisión", "Publicada"}
        );
        cmbEstado.setFont(new Font("SF Pro Text", Font.PLAIN, 16));
        cmbEstado.setBackground(CARD_BG);

        if (idTesis != -1) {
            cargarDatosTesis(
                    idTesis, txtNumero, txtTitulo, txtLicenciatura,
                    txtAutor, txtAsesor, cmbEstado, txtMes, txtAnio
            );
        }

        card.add(crearEtiqueta("Número*"));
        card.add(txtNumero);
        card.add(crearEtiqueta("Título*"));
        card.add(txtTitulo);
        card.add(crearEtiqueta("Licenciatura*"));
        card.add(txtLicenciatura);
        card.add(crearEtiqueta("Autor*"));
        card.add(txtAutor);
        card.add(crearEtiqueta("Asesor*"));
        card.add(txtAsesor);
        card.add(crearEtiqueta("Estado"));
        card.add(cmbEstado);
        card.add(crearEtiqueta("Mes"));
        card.add(txtMes);
        card.add(crearEtiqueta("Año"));
        card.add(txtAnio);

        JButton btnGuardar = crearBotonPrimario(
                idTesis == -1 ? "Guardar Tesis" : "Actualizar Tesis",
                e -> {
                    if (validarCamposNumericos(txtNumero, txtAnio)
                            && validarCamposTesis(txtNumero, txtTitulo, txtLicenciatura, txtAutor, txtAsesor)) {
                        if (idTesis == -1) {
                            guardarTesis(
                                    txtNumero.getText().trim(),
                                    txtTitulo.getText().trim(),
                                    txtLicenciatura.getText().trim(),
                                    txtAutor.getText().trim(),
                                    txtAsesor.getText().trim(),
                                    cmbEstado.getSelectedItem().toString(),
                                    txtMes.getText().trim(),
                                    txtAnio.getText().trim()
                            );
                        } else {
                            actualizarTesis(
                                    txtNumero.getText().trim(),
                                    txtTitulo.getText().trim(),
                                    txtLicenciatura.getText().trim(),
                                    txtAutor.getText().trim(),
                                    txtAsesor.getText().trim(),
                                    cmbEstado.getSelectedItem().toString(),
                                    txtMes.getText().trim(),
                                    txtAnio.getText().trim(),
                                    idTesis
                            );
                        }
                        dialog.dispose();
                    }
                }
        );

        card.add(Box.createVerticalStrut(20));
        card.add(btnGuardar);

        dialog.add(header, BorderLayout.NORTH);
        dialog.add(new JScrollPane(card), BorderLayout.CENTER);
        centrarVentana(dialog);
        dialog.setVisible(true);
    }

    private void cargarDatosTesis(
            int id,
            JTextField txtNumero,
            JTextField txtTitulo,
            JTextField txtLicenciatura,
            JTextField txtAutor,
            JTextField txtAsesor,
            JComboBox<String> cmbEstado,
            JTextField txtMes,
            JTextField txtAnio
    ) {
        String sql = "SELECT numero, titulo, licenciatura, autor, asesor, estado, mes, anio "
                + "FROM tesis WHERE id = ?";
        try (PreparedStatement stmt = conn.prepareStatement(sql)) {
            stmt.setInt(1, id);
            ResultSet rs = stmt.executeQuery();
            if (rs.next()) {
                txtNumero.setText(String.valueOf(rs.getInt("numero")));
                txtTitulo.setText(rs.getString("titulo"));
                txtLicenciatura.setText(rs.getString("licenciatura"));
                txtAutor.setText(rs.getString("autor"));
                txtAsesor.setText(rs.getString("asesor"));
                cmbEstado.setSelectedItem(rs.getString("estado"));
                txtMes.setText(rs.getString("mes"));
                Object a = rs.getObject("anio");
                txtAnio.setText(a == null ? "" : String.valueOf(a));
            }
        } catch (SQLException e) {
            mostrarError("Error al cargar datos: " + e.getMessage());
        }
    }

    private void guardarTesis(
            String numero,
            String titulo,
            String licenciatura,
            String autor,
            String asesor,
            String estado,
            String mes,
            String anio
    ) {
        try {
            int numTesis = Integer.parseInt(numero);
            int usuarioId = obtenerIdUsuario(currentUser);
            if (usuarioId == -1) {
                mostrarError("Usuario no encontrado");
                return;
            }

            String sql = "INSERT INTO tesis "
                    + "(numero,titulo,licenciatura,autor,asesor,estado,mes,anio,usuario_id) "
                    + "VALUES (?,?,?,?,?,?,?,?,?)";
            try (PreparedStatement stmt = conn.prepareStatement(sql)) {
                stmt.setInt(1, numTesis);
                stmt.setString(2, titulo);
                stmt.setString(3, licenciatura);
                stmt.setString(4, autor);
                stmt.setString(5, asesor);
                stmt.setString(6, estado);
                stmt.setString(7, mes.isEmpty() ? null : mes);
                stmt.setObject(8, anio.isEmpty() ? null : Integer.parseInt(anio));
                stmt.setInt(9, usuarioId);

                stmt.executeUpdate();
                cargarTesis("");
                mostrarMensaje("✅ Tesis registrada exitosamente");
            }
        } catch (NumberFormatException e) {
            mostrarError("Número y año deben ser valores numéricos");
        } catch (SQLException e) {
            mostrarError("Error al guardar: " + e.getMessage());
        }
    }

    private int obtenerIdUsuario(String usuario) {
        String sql = "SELECT id FROM usuarios WHERE usuario = ?";
        try (PreparedStatement stmt = conn.prepareStatement(sql)) {
            stmt.setString(1, usuario);
            ResultSet rs = stmt.executeQuery();
            return rs.next() ? rs.getInt("id") : -1;
        } catch (SQLException e) {
            return -1;
        }
    }

    private void actualizarTesis(
            String numero,
            String titulo,
            String licenciatura,
            String autor,
            String asesor,
            String estado,
            String mes,
            String anio,
            int id
    ) {
        try {
            int numTesis = Integer.parseInt(numero);

            String sql = "UPDATE tesis SET "
                    + "numero=?,titulo=?,licenciatura=?,autor=?,asesor=?,estado=?,mes=?,anio=? "
                    + "WHERE id = ?";
            try (PreparedStatement stmt = conn.prepareStatement(sql)) {
                stmt.setInt(1, numTesis);
                stmt.setString(2, titulo);
                stmt.setString(3, licenciatura);
                stmt.setString(4, autor);
                stmt.setString(5, asesor);
                stmt.setString(6, estado);
                stmt.setString(7, mes.isEmpty() ? null : mes);
                stmt.setObject(8, anio.isEmpty() ? null : Integer.parseInt(anio));
                stmt.setInt(9, id);

                int filas = stmt.executeUpdate();
                if (filas > 0) {
                    mostrarMensaje("✅ Tesis actualizada correctamente");
                    cargarTesis("");
                } else {
                    mostrarError("No se encontró la tesis");
                }
            }
        } catch (NumberFormatException e) {
            mostrarError("Número y año deben ser valores numéricos");
        } catch (SQLException e) {
            mostrarError("Error al actualizar: " + e.getMessage());
        }
    }

    private void eliminarTesis(int id) {
        String sql = "DELETE FROM tesis WHERE id = ?";
        try (PreparedStatement stmt = conn.prepareStatement(sql)) {
            stmt.setInt(1, id);
            int filas = stmt.executeUpdate();
            if (filas > 0) {
                mostrarMensaje("✅ Tesis eliminada");
                cargarTesis("");
            } else {
                mostrarError("No se encontró la tesis");
            }
        } catch (SQLException e) {
            mostrarError("Error al eliminar: " + e.getMessage());
        }
    }

    // ====== EXPORTAR A EXCEL ======
    private void exportarExcel() {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Guardar archivo Excel");
        chooser.setSelectedFile(new File("tesis_exportadas.xlsx"));
        if (chooser.showSaveDialog(currentFrame) != JFileChooser.APPROVE_OPTION) return;

        try (XSSFWorkbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(chooser.getSelectedFile())) {

            XSSFSheet sheet = workbook.createSheet("Tesis");

            // Estilo encabezado
            XSSFCellStyle headerStyle = workbook.createCellStyle();
            XSSFFont headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(new XSSFColor(new Color(200, 200, 200), null));
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Crear encabezados
            Row headerRow = sheet.createRow(0);
            for (int col = 0; col < tableModel.getColumnCount() - 1; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(tableModel.getColumnName(col));
                cell.setCellStyle(headerStyle);
            }

            // Escribir datos
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                Row row = sheet.createRow(i + 1);
                for (int j = 0; j < tableModel.getColumnCount() - 1; j++) {
                    Object val = tableModel.getValueAt(i, j);
                    Cell cell = row.createCell(j);
                    if (val instanceof Number) {
                        cell.setCellValue(((Number) val).doubleValue());
                    } else {
                        cell.setCellValue(val == null ? "" : val.toString());
                    }
                }
            }

            // Autoajustar columnas
            for (int col = 0; col < tableModel.getColumnCount() - 1; col++) {
                sheet.autoSizeColumn(col);
            }

            workbook.write(fos);
            mostrarMensaje("Exportación completada: " + chooser.getSelectedFile().getName());
        } catch (Exception ex) {
            mostrarError("Error exportando a Excel: " + ex.getMessage());
        }
    }

    // ====== COMPONENTES UI ======
    class ButtonRenderer extends JPanel implements TableCellRenderer {

        private final JButton btnEditar = new JButton("✎");
        private final JButton btnEliminar = new JButton("🗑");
        private final String currentUser;

        public ButtonRenderer(String currentUser) {
            this.currentUser = currentUser;
            setLayout(new FlowLayout(FlowLayout.CENTER, 5, 0));
            setBackground(CARD_BG);
            btnEditar.setBackground(COLOR_PRIMARIO);
            btnEditar.setForeground(Color.WHITE);
            btnEditar.setBorder(new RoundBorder(8, COLOR_PRIMARIO));
            btnEditar.setFocusPainted(false);
            btnEliminar.setBackground(COLOR_ELIMINAR);
            btnEliminar.setForeground(Color.WHITE);
            btnEliminar.setBorder(new RoundBorder(8, COLOR_ELIMINAR));
            btnEliminar.setFocusPainted(false);
            add(btnEditar);
            add(btnEliminar);
        }

        public Component getTableCellRendererComponent(
                JTable table, Object value, boolean isSelected,
                boolean hasFocus, int row, int column) {
            String userTesis = (String) table.getModel().getValueAt(row, 9);
            boolean ok = currentUser.equals(userTesis);
            btnEditar.setVisible(ok);
            btnEliminar.setVisible(ok);
            return this;
        }
    }

    class ButtonEditor extends DefaultCellEditor {

        private final JPanel panel = new JPanel(new FlowLayout(FlowLayout.CENTER, 5, 0));
        private final JButton btnEditar = new JButton("✎");
        private final JButton btnEliminar = new JButton("🗑");
        private final String currentUser;
        private int currentRow;

        public ButtonEditor(JCheckBox checkBox, String currentUser) {
            super(checkBox);
            this.currentUser = currentUser;
            panel.setBackground(CARD_BG);
            btnEditar.setBackground(COLOR_PRIMARIO);
            btnEditar.setForeground(Color.WHITE);
            btnEditar.setBorder(new RoundBorder(8, COLOR_PRIMARIO));
            btnEditar.setFocusPainted(false);
            btnEditar.addActionListener(e -> {
                int id = (int) tableModel.getValueAt(currentRow, 0);
                String userT = (String) tableModel.getValueAt(currentRow, 9);
                if (currentUser.equals(userT)) {
                    mostrarFormularioTesis(id, userT);
                } else {
                    mostrarError("Solo puedes editar tus propias tesis");
                }
                fireEditingStopped();
            });
            btnEliminar.setBackground(COLOR_ELIMINAR);
            btnEliminar.setForeground(Color.WHITE);
            btnEliminar.setBorder(new RoundBorder(8, COLOR_ELIMINAR));
            btnEliminar.setFocusPainted(false);
            btnEliminar.addActionListener(e -> {
                int id = (int) tableModel.getValueAt(currentRow, 0);
                String userT = (String) tableModel.getValueAt(currentRow, 9);
                if (currentUser.equals(userT)) {
                    mostrarDialogoEliminacion(id);
                } else {
                    mostrarError("Solo puedes eliminar tus propias tesis");
                }
                fireEditingStopped();
            });
            panel.add(btnEditar);
            panel.add(btnEliminar);
        }

        @Override
        public Component getTableCellEditorComponent(
                JTable table, Object value, boolean isSelected, int row, int column) {
            currentRow = row;
            String userT = (String) table.getModel().getValueAt(row, 9);
            boolean ok = currentUser.equals(userT);
            btnEditar.setVisible(ok);
            btnEliminar.setVisible(ok);
            return panel;
        }
    }

    private void mostrarDialogoEliminacion(int id) {
        JDialog dlg = new JDialog(currentFrame, "Confirmar eliminación", true);
        dlg.setSize(350, 200);
        dlg.setLayout(new BorderLayout());
        dlg.getContentPane().setBackground(FONDO);

        JPanel content = new JPanel(new BorderLayout());
        content.setBackground(FONDO);
        content.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));

        JLabel msg = new JLabel(
                "<html><div style='text-align:center;'>¿Estás seguro de eliminar esta tesis?<br>"
                + "<small>Esta acción no se puede deshacer</small></div>"
        );
        msg.setFont(new Font("SF Pro Text", Font.PLAIN, 16));
        msg.setHorizontalAlignment(SwingConstants.CENTER);

        JPanel btns = new JPanel(new FlowLayout(FlowLayout.CENTER, 15, 0));
        btns.setBackground(FONDO);
        JButton btnCanc = new JButton("Cancelar");
        btnCanc.addActionListener(e -> dlg.dispose());
        JButton btnOk = new JButton("Eliminar");
        btnOk.setBackground(COLOR_ELIMINAR);
        btnOk.setForeground(Color.WHITE);
        btnOk.addActionListener(e -> {
            eliminarTesis(id);
            dlg.dispose();
        });
        btns.add(btnCanc);
        btns.add(btnOk);

        content.add(msg, BorderLayout.CENTER);
        content.add(btns, BorderLayout.SOUTH);

        dlg.add(content, BorderLayout.CENTER);
        centrarVentana(dlg);
        dlg.setVisible(true);
    }

    // ====== MÉTODOS AUXILIARES ======
    private JPanel crearPanelHeader(String t, String st) {
        JPanel header = new JPanel();
        header.setLayout(new BoxLayout(header, BoxLayout.Y_AXIS));
        header.setBackground(FONDO);
        header.setBorder(BorderFactory.createEmptyBorder(40, 0, 20, 0));
        JLabel lt = new JLabel(t);
        lt.setFont(new Font("SF Pro Display", Font.BOLD, 28));
        lt.setAlignmentX(Component.CENTER_ALIGNMENT);
        JLabel ls = new JLabel(st);
        ls.setFont(new Font("SF Pro Text", Font.PLAIN, 14));
        ls.setForeground(Color.GRAY);
        ls.setAlignmentX(Component.CENTER_ALIGNMENT);
        header.add(lt);
        header.add(Box.createVerticalStrut(10));
        header.add(ls);
        return header;
    }

    private JTextField crearCampoTexto(String placeholder) {
        JTextField campo = new JTextField();
        campo.setFont(new Font("SF Pro Text", Font.PLAIN, 16));
        campo.setBorder(BorderFactory.createCompoundBorder(
                new RoundBorder(8, new Color(200, 200, 200)),
                BorderFactory.createEmptyBorder(12, 15, 12, 15)
        ));
        campo.putClientProperty("JTextField.placeholder", placeholder);
        return campo;
    }

    private JPasswordField crearCampoTextoContrasena() {
        JPasswordField campo = new JPasswordField();
        campo.setFont(new Font("SF Pro Text", Font.PLAIN, 16));
        campo.setBorder(BorderFactory.createCompoundBorder(
                new RoundBorder(8, new Color(200, 200, 200)),
                BorderFactory.createEmptyBorder(12, 15, 12, 15)
        ));
        campo.putClientProperty("JTextField.placeholder", "Contraseña");
        return campo;
    }

    private JLabel crearEtiqueta(String texto) {
        JLabel lbl = new JLabel(texto);
        lbl.setFont(new Font("SF Pro Text", Font.PLAIN, 14));
        lbl.setForeground(new Color(60, 60, 67));
        return lbl;
    }

    private JButton crearBotonPrimario(String texto, ActionListener accion) {
        JButton btn = new JButton(texto);
        btn.setFont(new Font("SF Pro Text", Font.BOLD, 16));
        btn.setBackground(COLOR_PRIMARIO);
        btn.setForeground(Color.WHITE);
        btn.setFocusPainted(false);
        btn.setBorder(new RoundBorder(12, COLOR_PRIMARIO));
        btn.setCursor(new Cursor(Cursor.HAND_CURSOR));
        if (accion != null) btn.addActionListener(accion);
        return btn;
    }

    private JButton crearBotonSecundario(String texto, ActionListener accion) {
        JButton btn = new JButton(texto);
        btn.setFont(new Font("SF Pro Text", Font.PLAIN, 14));
        btn.setForeground(COLOR_PRIMARIO);
        btn.setContentAreaFilled(false);
        btn.setBorderPainted(false);
        btn.setCursor(new Cursor(Cursor.HAND_CURSOR));
        if (accion != null) btn.addActionListener(accion);
        return btn;
    }

    private JButton crearBotonIcono(String icono) {
        JButton btn = new JButton(icono);
        btn.setFont(new Font("SF Pro Text", Font.PLAIN, 18));
        btn.setBackground(FONDO);
        btn.setBorder(new RoundBorder(8, new Color(200, 200, 200)));
        btn.setCursor(new Cursor(Cursor.HAND_CURSOR));
        return btn;
    }

    private void estilizarTabla(JTable tabla) {
        tabla.setFont(new Font("SF Pro Text", Font.PLAIN, 14));
        tabla.setRowHeight(40);
        tabla.setShowVerticalLines(false);
        tabla.setIntercellSpacing(new Dimension(0, 5));
        JTableHeader h = tabla.getTableHeader();
        h.setFont(new Font("SF Pro Text", Font.BOLD, 14));
        h.setBackground(CARD_BG);
        h.setForeground(new Color(60, 60, 67));
        h.setReorderingAllowed(false);
        tabla.setDefaultRenderer(Object.class, new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(
                    JTable table, Object value, boolean selected,
                    boolean focus, int row, int col) {
                JLabel c = (JLabel) super.getTableCellRendererComponent(
                        table, value, selected, focus, row, col);
                if (col == 6) {
                    String est = value.toString();
                    Color color;
                    switch (est) {
                        case "Finalizada":
                            color = new Color(52, 199, 89);
                            break;
                        case "En revisión":
                            color = new Color(255, 149, 0);
                            break;
                        case "Publicada":
                            color = new Color(88, 86, 214);
                            break;
                        default:
                            color = new Color(142, 142, 147);
                    }
                    c.setForeground(color);
                    c.setHorizontalAlignment(SwingConstants.CENTER);
                }
                if (col == 9) {
                    c.setHorizontalAlignment(SwingConstants.CENTER);
                }
                return c;
            }
        });
    }
    private boolean validarCamposTesis(JTextField... campos) {
        for (JTextField campo : campos) {
            if (campo.getText().trim().isEmpty()) {
                mostrarError("Los campos marcados con (*) son obligatorios");
                campo.setBorder(BorderFactory.createLineBorder(Color.RED, 1));
                campo.grabFocus();
                return false;
            }
            campo.setBorder(new RoundBorder(8, new Color(200, 200, 200)));
        }
        return true;
    }

    private boolean validarCamposNumericos(JTextField... campos) {
        try {
            for (JTextField campo : campos) {
                if (!campo.getText().trim().isEmpty()) {
                    Integer.parseInt(campo.getText().trim());
                }
            }
            return true;
        } catch (NumberFormatException e) {
            mostrarError("Los campos numéricos deben contener valores válidos");
            return false;
        }
    }

    private void centrarVentana(Window w) {
        w.setLocationRelativeTo(null);
    }

    private void mostrarMensaje(String msg) {
        JOptionPane.showMessageDialog(currentFrame, msg, "Tesla Tesis", JOptionPane.INFORMATION_MESSAGE);
    }

    private void mostrarError(String msg) {
        JOptionPane.showMessageDialog(currentFrame, msg, "Error", JOptionPane.ERROR_MESSAGE);
    }

    class RoundBorder extends AbstractBorder {

        private final int radius;
        private final Color color;

        public RoundBorder(int radius, Color color) {
            this.radius = radius;
            this.color = color;
        }

        @Override
        public void paintBorder(Component c, Graphics g, int x, int y, int w, int h) {
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(
                    RenderingHints.KEY_ANTIALIASING,
                    RenderingHints.VALUE_ANTIALIAS_ON
            );
            g2.setColor(color);
            g2.drawRoundRect(x, y, w - 1, h - 1, radius, radius);
            g2.dispose();
        }

        @Override
        public Insets getBorderInsets(Component c) {
            return new Insets(radius + 1, radius + 1, radius + 1, radius + 1);
        }

        @Override
        public Insets getBorderInsets(Component c, Insets insets) {
            insets.set(radius + 1, radius + 1, radius + 1, radius + 1);
            return insets;
        }
    }
}
