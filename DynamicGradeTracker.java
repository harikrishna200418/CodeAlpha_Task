import javafx.animation.Interpolator;
import javafx.animation.ScaleTransition;
import javafx.animation.TranslateTransition;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.input.KeyCode;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Duration;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
import java.io.*;
import java.net.URL;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class DynamicGradeTracker extends Application {

    private final TableView<ObservableList<SimpleStringProperty>> tableView = new TableView<>();
    private final ObservableList<ObservableList<SimpleStringProperty>> data = FXCollections.observableArrayList();
    private List<String> subjects = new ArrayList<>();
    private ObservableList<String> maxMarks = FXCollections.observableArrayList(); // Stores max marks for each subject
    private final Label overallClassAverageLabel = new Label("N/A");
    private final Label summaryTitleLabel = new Label("Overall Class Average");
    private final TextField subjectNamesField = new TextField();
    private final TextField maxMarksField = new TextField();
    private final ScriptEngine scriptEngine = new ScriptEngineManager().getEngineByName("JavaScript");
    private static final String AUTOSAVE_FILE = "grades_autosave.csv";

    // New: Map column characters to data indices for formula parsing (A=1, B=2
    // etc.)
    private final Map<Character, Integer> columnIndexMap = new HashMap<>();
    // New: Regex pattern to find cell references (e.g., A1, C23)
    private static final Pattern CELL_REF_PATTERN = Pattern.compile("([A-Z]+)(\\d+)");
    // New: Regex pattern to find SUM/AVERAGE functions with comma-separated args or
    // range
    private static final Pattern FUNCTION_PATTERN = Pattern
            .compile("(SUM|AVERAGE)\\(([A-Z]+)(\\d+)(?::([A-Z]+)(\\d+))?((?:,[A-Z]+\\d+)*)\\)");

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Dynamic Grade Tracker ✨");

        BorderPane root = new BorderPane();
        root.getStyleClass().add("root-pane");

        VBox controlPanel = createControlPanel();
        root.setLeft(controlPanel);

        VBox rightPanel = createRightPanel();
        root.setCenter(rightPanel);

        initializeColumnMaps(); // Initialize column letter to index mapping
        loadAutoSavedData(); // Attempt to load data at startup

        Scene scene = new Scene(root, 1400, 850);
        URL cssUrl = getClass().getResource("styles.css");
        if (cssUrl != null) {
            scene.getStylesheets().add(cssUrl.toExternalForm());
        } else {
            System.err.println("Warning: Could not find 'styles.css'. Application will use default styling.");
        }
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    // New: Initialize column character to index mapping for formula parsing
    private void initializeColumnMaps() {
        // Fixed columns: S.NO (internal 0), ROLL NO (1), NAME (2)
        // Subject columns start from index 3
        columnIndexMap.put('A', 1); // ROLL NO
        columnIndexMap.put('B', 2); // NAME

        for (int i = 0; i < 26; i++) { // For A-Z
            char colChar = (char) ('C' + i);
            columnIndexMap.put(colChar, i + 3); // C maps to 3, D to 4, etc.
        }
        // Could extend for AA, AB etc. but 26 subjects is a lot already.
    }

    private int getColumnIndexFromChar(String colChars) {
        int index = 0;
        for (char c : colChars.toCharArray()) {
            index = index * 26 + (c - 'A' + 1);
        }
        // Adjust for 0-based indexing and fixed columns offset (A->1, B->2, C->3 etc.
        // mapping to data indices)
        return index; // This needs careful mapping if you use AA, AB etc. For now, use the map
    }

    private VBox createControlPanel() {
        VBox controlPanel = new VBox(30);
        controlPanel.setPadding(new Insets(25));
        controlPanel.getStyleClass().add("control-panel");
        controlPanel.setPrefWidth(340);

        // Section 1: Define Subjects & Max Marks
        VBox defineSubjectsContent = new VBox(15);
        Label subjectNamesLabel = new Label("Subject Names (comma-separated):");
        subjectNamesField.setPromptText("e.g., Maths, Science, History");

        Label maxMarksLabel = new Label("Max Marks per Subject (comma-separated):");
        maxMarksField.setPromptText("e.g., 100, 50, 80 (must match subjects)");

        Button updateSubjectsButton = new Button("🔄 Update Subjects");
        updateSubjectsButton.setMaxWidth(Double.MAX_VALUE);
        updateSubjectsButton.getStyleClass().add("button-primary");
        updateSubjectsButton.setOnAction(e -> updateSubjects(subjectNamesField.getText(), maxMarksField.getText()));

        defineSubjectsContent.getChildren().addAll(subjectNamesLabel, subjectNamesField, maxMarksLabel, maxMarksField,
                updateSubjectsButton);
        TitledPane defineSubjectsPane = new TitledPane("1. Define Subjects & Max Marks", defineSubjectsContent);
        defineSubjectsPane.setCollapsible(false);

        // Section 2: Create Student Rows
        VBox createStudentRowsContent = new VBox(15);
        Label numStudentsLabel = new Label("Number of Students to Add:");
        TextField numStudentsField = new TextField();
        numStudentsField.setPromptText("e.g., 5");
        Button createStudentRowsButton = new Button("➕ Create Student Rows");
        createStudentRowsButton.setMaxWidth(Double.MAX_VALUE);
        createStudentRowsButton.getStyleClass().add("button-success");
        createStudentRowsButton.setOnAction(e -> createStudentRows(numStudentsField.getText()));
        createStudentRowsContent.getChildren().addAll(numStudentsLabel, numStudentsField, createStudentRowsButton);
        TitledPane createStudentRowsPane = new TitledPane("2. Create Student Rows", createStudentRowsContent);
        createStudentRowsPane.setCollapsible(false);

        // Calculate All Student Stats (and class description)
        Button calculateAllStatsButton = new Button("🧮 Calculate All Student Stats");
        calculateAllStatsButton.setMaxWidth(Double.MAX_VALUE);
        calculateAllStatsButton.getStyleClass().add("button-primary");
        calculateAllStatsButton.setOnAction(e -> recalculateAllStudentStats());

        // Action Buttons
        Button importButton = new Button("📥 Import from CSV");
        importButton.getStyleClass().add("button-info");
        importButton.setMaxWidth(Double.MAX_VALUE);
        importButton.setOnAction(e -> importFromCSV());

        Button exportButton = new Button("💾 Export to CSV");
        exportButton.getStyleClass().add("button-info");
        exportButton.setMaxWidth(Double.MAX_VALUE);
        exportButton.setOnAction(e -> exportToCSV());

        Button cleanupButton = new Button("🧹 Clean Empty Rows");
        cleanupButton.getStyleClass().add("button-warning");
        cleanupButton.setMaxWidth(Double.MAX_VALUE);
        cleanupButton.setOnAction(e -> handleManualCleanup());

        Button deleteSelectedButton = new Button("🗑️ Delete Selected Rows");
        deleteSelectedButton.getStyleClass().add("button-danger");
        deleteSelectedButton.setMaxWidth(Double.MAX_VALUE);
        deleteSelectedButton.setOnAction(e -> handleDeleteSelectedRows());

        Button clearAllButton = new Button("💥 Clear All Data");
        clearAllButton.getStyleClass().add("button-danger-outline");
        clearAllButton.setMaxWidth(Double.MAX_VALUE);
        clearAllButton.setOnAction(e -> clearAllData());

        VBox bottomButtons = new VBox(15, calculateAllStatsButton, importButton, exportButton, cleanupButton,
                deleteSelectedButton,
                clearAllButton);
        bottomButtons.setAlignment(Pos.CENTER);

        controlPanel.getChildren().addAll(defineSubjectsPane, createStudentRowsContent, new Region(), bottomButtons);
        VBox.setVgrow(controlPanel.getChildren().get(2), Priority.ALWAYS);

        return controlPanel;
    }

    private VBox createRightPanel() {
        VBox rightPanel = new VBox(20);
        rightPanel.setPadding(new Insets(25));
        rightPanel.getStyleClass().add("right-panel");

        Label titleLabel = new Label("Student Grade Roster");
        titleLabel.getStyleClass().add("header-title");

        Label subtitleLabel = new Label(
                "Import a CSV, or define subjects to begin. Use Ctrl/Cmd+Click to select rows. Formulas start with '=' (e.g., =50*2, =SUM(C2:C4))"); // Updated
                                                                                                                                                     // subtitle
        subtitleLabel.getStyleClass().add("header-subtitle");

        tableView.setEditable(true);
        tableView.setItems(data);
        tableView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        VBox.setVgrow(tableView, Priority.ALWAYS);

        Label placeholder = new Label("Import a CSV or define subjects and add students to begin! 🚀");
        placeholder.getStyleClass().add("table-placeholder");
        tableView.setPlaceholder(placeholder);

        tableView.setRowFactory(tv -> {
            TableRow<ObservableList<SimpleStringProperty>> row = new TableRow<>();
            ContextMenu contextMenu = new ContextMenu();
            MenuItem addAbove = new MenuItem("⬆️ Add Row Above");
            addAbove.setOnAction(e -> addRowAt(row.getIndex()));
            MenuItem addBelow = new MenuItem("⬇️ Add Row Below");
            addBelow.setOnAction(e -> addRowAt(row.getIndex() + 1));
            MenuItem deleteRow = new MenuItem("❌ Delete This Row");
            deleteRow.getStyleClass().add("menu-item-danger");
            deleteRow.setOnAction(e -> {
                data.remove(row.getItem());
                autoSaveData();
            });
            contextMenu.getItems().addAll(addAbove, addBelow, new SeparatorMenuItem(), deleteRow);
            row.contextMenuProperty().bind(row.emptyProperty().map(empty -> empty ? null : contextMenu));

            // Add hover effect to rows
            row.hoverProperty().addListener((obs, wasHovered, isNowHovered) -> {
                if (isNowHovered) {
                    row.getStyleClass().add("table-row-cell-hover");
                } else {
                    row.getStyleClass().remove("table-row-cell-hover");
                }
            });
            return row;
        });

        updateTableColumns();

        VBox summaryCard = new VBox(5, summaryTitleLabel, overallClassAverageLabel);
        summaryCard.getStyleClass().add("summary-card");
        summaryCard.setAlignment(Pos.CENTER);
        summaryTitleLabel.getStyleClass().add("summary-title");
        overallClassAverageLabel.getStyleClass().add("summary-value");

        Button generateSummaryButton = new Button("📊 Generate Class Summary");
        generateSummaryButton.getStyleClass().add("button-primary");
        generateSummaryButton.setOnAction(e -> {
            calculateOverallAverage();
            ScaleTransition st = new ScaleTransition(Duration.millis(200), summaryCard);
            st.setFromX(1);
            st.setFromY(1);
            st.setToX(1.05);
            st.setToY(1.05);
            st.setCycleCount(2);
            st.setAutoReverse(true);
            st.play();
        });

        HBox bottomBar = new HBox(30, summaryCard, generateSummaryButton);
        bottomBar.setAlignment(Pos.CENTER_RIGHT);
        HBox.setHgrow(summaryCard, Priority.ALWAYS);

        rightPanel.getChildren().addAll(titleLabel, subtitleLabel, tableView, bottomBar);
        return rightPanel;
    }

    private void updateSubjects(String subjectNamesText, String maxMarksText) {
        this.subjects = new ArrayList<>(Arrays.asList(subjectNamesText.split("\\s*,\\s*")));
        if (this.subjects.size() == 1 && this.subjects.get(0).isEmpty()) {
            this.subjects.clear();
        }

        List<String> tempMaxMarks = Arrays.asList(maxMarksText.split("\\s*,\\s*"));
        this.maxMarks.clear();

        if (this.subjects.size() == tempMaxMarks.size()) {
            this.maxMarks.addAll(tempMaxMarks);
        } else if (!tempMaxMarks.isEmpty() && !(tempMaxMarks.size() == 1 && tempMaxMarks.get(0).isEmpty())) {
            showAlert("Input Error", "Number of 'Max Marks' entries (" + tempMaxMarks.size() +
                    ") must match the number of 'Subjects' (" + this.subjects.size() +
                    "). Defaulting to 100 for all subjects.");
            for (int i = 0; i < this.subjects.size(); i++) {
                this.maxMarks.add("100");
            }
            maxMarksField.setText(String.join(", ", this.maxMarks));
        } else {
            for (int i = 0; i < this.subjects.size(); i++) {
                this.maxMarks.add("100");
            }
            maxMarksField.setText(String.join(", ", this.maxMarks));
        }

        updateTableColumns();
        adjustDataToNewColumns();
        autoSaveData();
    }

    private void updateTableColumns() {
        tableView.getColumns().clear();

        TableColumn<ObservableList<SimpleStringProperty>, String> snoColumn = new TableColumn<>("S.NO.");
        snoColumn.setPrefWidth(60);
        snoColumn.setSortable(false);
        snoColumn.setEditable(false);
        snoColumn.getStyleClass().add("centered-cell");
        snoColumn.setCellValueFactory(
                param -> new SimpleStringProperty(String.valueOf(data.indexOf(param.getValue()) + 1)));

        TableColumn<ObservableList<SimpleStringProperty>, String> rollNoColumn = createEditableColumn("ROLL NO", 1);
        rollNoColumn.getStyleClass().add("centered-cell");

        TableColumn<ObservableList<SimpleStringProperty>, String> nameColumn = createEditableColumn("NAME", 2);
        nameColumn.setPrefWidth(150);

        tableView.getColumns().addAll(snoColumn, rollNoColumn, nameColumn);

        // Update columnIndexMap for dynamic subject columns
        char currentSubjectColChar = 'C'; // First subject is column 'C'
        for (int i = 0; i < subjects.size(); i++) {
            TableColumn<ObservableList<SimpleStringProperty>, String> subjectColumn = createEditableColumn(
                    subjects.get(i).toUpperCase(), i + 3);
            subjectColumn.getStyleClass().add("centered-cell");
            tableView.getColumns().add(subjectColumn);
            columnIndexMap.put(currentSubjectColChar, i + 3); // Map 'C'->3, 'D'->4 etc.
            currentSubjectColChar++;
        }

        TableColumn<ObservableList<SimpleStringProperty>, String> totalMarksColumn = new TableColumn<>("TOTAL MARKS");
        totalMarksColumn.setPrefWidth(120);
        totalMarksColumn.setSortable(false);
        totalMarksColumn.setEditable(false);
        totalMarksColumn.getStyleClass().add("centered-cell");
        totalMarksColumn.setCellValueFactory(
                param -> new SimpleStringProperty(String.format("%.2f", calculateRowTotalMarks(param.getValue()))));
        tableView.getColumns().add(totalMarksColumn);

        TableColumn<ObservableList<SimpleStringProperty>, String> percentageColumn = new TableColumn<>("PERCENTAGE");
        percentageColumn.setPrefWidth(120);
        percentageColumn.setSortable(true);
        percentageColumn.setEditable(false);
        percentageColumn.getStyleClass().add("centered-cell");
        percentageColumn.setCellValueFactory(
                param -> new SimpleStringProperty(String.format("%.2f%%", calculateRowPercentage(param.getValue()))));
        tableView.getColumns().add(percentageColumn);

        TableColumn<ObservableList<SimpleStringProperty>, String> gradeColumn = new TableColumn<>("GRADE");
        gradeColumn.setPrefWidth(90);
        gradeColumn.setSortable(true);
        gradeColumn.setEditable(false);
        gradeColumn.getStyleClass().add("centered-cell");
        gradeColumn.setCellValueFactory(
                param -> new SimpleStringProperty(calculateGrade(calculateRowPercentage(param.getValue()))));
        gradeColumn.setCellFactory(column -> new GradeCell());
        tableView.getColumns().add(gradeColumn);

        // New columns for Average, Highest, and Lowest marks
        TableColumn<ObservableList<SimpleStringProperty>, String> avgMarkColumn = new TableColumn<>("AVG MARK");
        avgMarkColumn.setPrefWidth(120);
        avgMarkColumn.setSortable(false);
        avgMarkColumn.setEditable(false);
        avgMarkColumn.getStyleClass().add("centered-cell");
        avgMarkColumn.setCellValueFactory(
                param -> new SimpleStringProperty(String.format("%.2f", calculateAverageMark(param.getValue()))));
        tableView.getColumns().add(avgMarkColumn);

        TableColumn<ObservableList<SimpleStringProperty>, String> highestMarkColumn = new TableColumn<>("HIGHEST MARK");
        highestMarkColumn.setPrefWidth(120);
        highestMarkColumn.setSortable(false);
        highestMarkColumn.setEditable(false);
        highestMarkColumn.getStyleClass().add("centered-cell");
        highestMarkColumn.setCellValueFactory(
                param -> new SimpleStringProperty(String.format("%.2f", calculateHighestMark(param.getValue()))));
        tableView.getColumns().add(highestMarkColumn);

        TableColumn<ObservableList<SimpleStringProperty>, String> lowestMarkColumn = new TableColumn<>("LOWEST MARK");
        lowestMarkColumn.setPrefWidth(120);
        lowestMarkColumn.setSortable(false);
        lowestMarkColumn.setEditable(false);
        lowestMarkColumn.getStyleClass().add("centered-cell");
        lowestMarkColumn.setCellValueFactory(
                param -> new SimpleStringProperty(String.format("%.2f", calculateLowestMark(param.getValue()))));
        tableView.getColumns().add(lowestMarkColumn);
    }

    private TableColumn<ObservableList<SimpleStringProperty>, String> createEditableColumn(String title, int index) {
        TableColumn<ObservableList<SimpleStringProperty>, String> column = new TableColumn<>(title);
        column.setPrefWidth(120);
        column.setCellValueFactory(param -> {
            while (param.getValue().size() <= index) {
                param.getValue().add(new SimpleStringProperty(""));
            }
            return param.getValue().get(index);
        });
        column.setCellFactory(col -> new AlwaysEditingCell());
        column.setOnEditCommit(event -> {
            event.getRowValue().get(index).set(event.getNewValue());
            // When any cell is edited, refresh the entire table to ensure all dependent
            // calculations update.
            // This is a simple (though not Excel-efficient) way to handle dependencies.
            tableView.refresh();
            autoSaveData();
        });
        return column;
    }

    private void createStudentRows(String numStudentsStr) {
        try {
            int numStudents = Integer.parseInt(numStudentsStr);
            if (numStudents <= 0) {
                showAlert("Input Error", "Please enter a positive number of students.");
                return;
            }
            for (int i = 0; i < numStudents; i++) {
                data.add(createRow(new String[2 + subjects.size()]));
            }
            autoSaveData();
        } catch (NumberFormatException e) {
            showAlert("Input Error", "Please enter a valid number.");
        }
    }

    private void adjustDataToNewColumns() {
        int newCoreColumnCount = 1 + 2 + subjects.size();
        for (ObservableList<SimpleStringProperty> row : data) {
            while (row.size() < newCoreColumnCount) {
                row.add(new SimpleStringProperty(""));
            }
            while (row.size() > newCoreColumnCount) {
                if (row.size() > 3) {
                    row.remove(row.size() - 1);
                } else {
                    break;
                }
            }
        }
        tableView.refresh();
    }

    private void calculateOverallAverage() {
        double totalOverallPercentage = 0;
        int validStudentCount = 0;
        boolean hasErrors = false;

        for (ObservableList<SimpleStringProperty> row : data) {
            double studentPercentage = calculateRowPercentage(row);
            if (!Double.isNaN(studentPercentage) && !Double.isInfinite(studentPercentage)) {
                totalOverallPercentage += studentPercentage;
                validStudentCount++;
            } else {
                hasErrors = true;
            }
        }

        var summaryCard = summaryTitleLabel.getParent();
        summaryCard.getStyleClass().remove("summary-error");

        if (hasErrors || validStudentCount == 0) {
            overallClassAverageLabel.setText("Error");
            summaryCard.getStyleClass().add("summary-error");
        } else {
            double overallAverage = (totalOverallPercentage / validStudentCount);
            overallClassAverageLabel.setText(String.format("%.2f%%", overallAverage));
        }
    }

    private void recalculateAllStudentStats() {
        tableView.refresh(); // Recalculates individual student stats
        calculateOverallAverage(); // Recalculates and updates the overall class average
        showAlert("Recalculation Complete",
                "All student total marks, percentages, and grades have been recalculated. Overall class summary updated.");
    }

    private void writeToCsv(Writer writer) throws IOException {
        writer.append(String.join(",", subjects)).append("\n");
        writer.append(maxMarks.stream().collect(Collectors.joining(","))).append("\n");

        for (var row : data) {
            String record = row.stream()
                    .skip(1)
                    .limit(2 + subjects.size())
                    .map(SimpleStringProperty::get)
                    .collect(Collectors.joining(","));
            writer.append(record).append("\n");
        }
    }

    private void autoSaveData() {
        data.removeIf(this::isRowEmpty);
        File file = new File(AUTOSAVE_FILE);
        try (FileWriter writer = new FileWriter(file)) {
            writeToCsv(writer);
        } catch (IOException e) {
            System.err.println("Auto-save failed: " + e.getMessage());
        }
    }

    private void loadAutoSavedData() {
        File file = new File(AUTOSAVE_FILE);
        if (!file.exists()) {
            return;
        }
        try (BufferedReader reader = new BufferedReader(new FileReader(file))) {
            importDataFromReader(reader, null);
        } catch (IOException e) {
            System.err.println("Failed to load auto-saved data: " + e.getMessage());
        }
    }

    private void importFromCSV() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Import from CSV");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("CSV Files", "*.csv"));
        File file = fileChooser.showOpenDialog(null);
        if (file == null) {
            return;
        }
        try (BufferedReader reader = new BufferedReader(new FileReader(file))) {
            importDataFromReader(reader, "Successfully imported data from " + file.getName());
        } catch (IOException e) {
            showAlert("Import Error", "Failed to read the file. Please ensure it is a valid CSV file.");
            e.printStackTrace();
        }
    }

    private void exportToCSV() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Save as CSV");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("CSV Files", "*.csv"));
        File file = fileChooser.showSaveDialog(null);
        if (file != null) {
            try (FileWriter writer = new FileWriter(file)) {
                writeToCsv(writer);
                showAlert("Success", "Data exported successfully to " + file.getName());
            } catch (IOException e) {
                showAlert("Error", "Failed to export data to CSV.");
                e.printStackTrace();
            }
        }
    }

    private void importDataFromReader(BufferedReader reader, String successMessage) throws IOException {
        String subjectsLine = reader.readLine();
        String maxMarksLine = reader.readLine();

        if (subjectsLine == null || maxMarksLine == null) {
            if (successMessage != null) {
                showAlert("Import Error",
                        "The selected CSV file is empty or malformed (missing subject/max marks header).");
            }
            return;
        }

        List<ObservableList<SimpleStringProperty>> importedData = new ArrayList<>();
        String dataLine;
        while ((dataLine = reader.readLine()) != null) {
            String[] values = dataLine.split(",", -1);
            if (values.length < 2) {
                continue;
            }
            importedData.add(createRow(values));
        }

        Platform.runLater(() -> {
            subjectNamesField.setText(subjectsLine);
            maxMarksField.setText(maxMarksLine);
            updateSubjects(subjectsLine, maxMarksLine);
            data.setAll(importedData);
            if (successMessage != null) {
                showAlert("Success", successMessage);
            }
        });
    }

    private void clearAllData() {
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setTitle("Confirm Clear");
        alert.setHeaderText("Clear All Data");
        alert.setContentText("Are you sure you want to permanently delete all subjects and student data?");
        alert.getDialogPane().getStyleClass().add("custom-dialog");
        Optional<ButtonType> result = alert.showAndWait();
        if (result.isPresent() && result.get() == ButtonType.OK) {
            subjects.clear();
            maxMarks.clear();
            data.clear();
            subjectNamesField.setText("");
            maxMarksField.setText("");
            updateSubjects("", "");
            overallClassAverageLabel.setText("N/A");
            new File(AUTOSAVE_FILE).delete();
        }
    }

    private void showAlert(String title, String message) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.getDialogPane().getStyleClass().add("custom-dialog");
        alert.showAndWait();
    }

    private void addRowAt(int index) {
        String[] emptyValues = new String[2 + subjects.size()];
        Arrays.fill(emptyValues, "");
        data.add(index, createRow(emptyValues));
        autoSaveData();
    }

    private void handleDeleteSelectedRows() {
        List<ObservableList<SimpleStringProperty>> selectedRows = new ArrayList<>(
                tableView.getSelectionModel().getSelectedItems());
        if (selectedRows.isEmpty()) {
            showAlert("No Selection", "Please select one or more rows to delete.");
            return;
        }
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setTitle("Confirm Deletion");
        alert.setHeaderText("Delete Selected Rows");
        alert.setContentText(
                "Are you sure you want to permanently delete the " + selectedRows.size() + " selected row(s)?");
        alert.getDialogPane().getStyleClass().add("custom-dialog");
        Optional<ButtonType> result = alert.showAndWait();
        if (result.isPresent() && result.get() == ButtonType.OK) {
            data.removeAll(selectedRows);
            autoSaveData();
        }
    }

    private void handleManualCleanup() {
        int initialSize = data.size();
        data.removeIf(this::isRowEmpty);
        int rowsRemoved = initialSize - data.size();
        if (rowsRemoved > 0) {
            showAlert("Cleanup Complete", "Removed " + rowsRemoved + " empty row(s).");
            autoSaveData();
        } else {
            showAlert("No Empty Rows", "No empty rows were found to clean up.");
        }
    }

    private boolean isRowEmpty(ObservableList<SimpleStringProperty> row) {
        return row.stream().skip(1).allMatch(p -> p.get().trim().isEmpty());
    }

    private ObservableList<SimpleStringProperty> createRow(String... values) {
        List<SimpleStringProperty> list = new ArrayList<>();
        list.add(new SimpleStringProperty("")); // For S.NO (internal)

        for (String value : values) {
            list.add(new SimpleStringProperty(value));
        }

        int requiredEditableColumns = 2 + subjects.size();
        while (list.size() - 1 < requiredEditableColumns) {
            list.add(new SimpleStringProperty(""));
        }
        return FXCollections.observableArrayList(list);
    }

    // New: Helper method to get a list of valid numerical marks for a student row.
    private List<Double> getStudentMarks(ObservableList<SimpleStringProperty> row) {
        List<Double> marks = new ArrayList<>();
        for (int i = 0; i < subjects.size(); i++) {
            try {
                String scoreStr = getProcessedValue(row.get(i + 3).get(), data.indexOf(row)); // Subjects start at index
                                                                                              // 3
                if (scoreStr != null && !scoreStr.isEmpty() && !scoreStr.contains("Error")) {
                    marks.add(Double.parseDouble(scoreStr));
                }
            } catch (IndexOutOfBoundsException | NumberFormatException e) {
                // Ignore if cell doesn't exist or isn't a valid number
            }
        }
        return marks;
    }

    // New: Calculates the average of marks obtained by a student.
    private double calculateAverageMark(ObservableList<SimpleStringProperty> row) {
        List<Double> marks = getStudentMarks(row);
        if (marks.isEmpty()) {
            return 0.0;
        }
        return marks.stream().mapToDouble(Double::doubleValue).average().orElse(0.0);
    }

    // New: Finds the highest mark obtained by a student.
    private double calculateHighestMark(ObservableList<SimpleStringProperty> row) {
        List<Double> marks = getStudentMarks(row);
        if (marks.isEmpty()) {
            return 0.0;
        }
        return Collections.max(marks);
    }

    // New: Finds the lowest mark obtained by a student.
    private double calculateLowestMark(ObservableList<SimpleStringProperty> row) {
        List<Double> marks = getStudentMarks(row);
        if (marks.isEmpty()) {
            return 0.0;
        }
        return Collections.min(marks);
    }

    private double calculateRowTotalMarks(ObservableList<SimpleStringProperty> row) {
        return getStudentMarks(row).stream().mapToDouble(Double::doubleValue).sum();
    }

    private double calculateRowPercentage(ObservableList<SimpleStringProperty> row) {
        double totalObtainedMarks = calculateRowTotalMarks(row);
        double totalMaxMarks = 0;

        for (int i = 0; i < subjects.size(); i++) {
            try {
                // We add max marks only if the student has a corresponding valid mark
                // to avoid penalizing students for subjects they haven't been marked for yet.
                String scoreStr = getProcessedValue(row.get(i + 3).get(), data.indexOf(row));
                if (scoreStr != null && !scoreStr.isEmpty() && !scoreStr.contains("Error")) {
                    double currentMaxMark = (i < maxMarks.size() && !maxMarks.get(i).trim().isEmpty())
                            ? Double.parseDouble(maxMarks.get(i))
                            : 100;
                    totalMaxMarks += currentMaxMark;
                }
            } catch (IndexOutOfBoundsException | NumberFormatException e) {
                // Ignore if max mark is invalid or student mark doesn't exist.
            }
        }
        return (totalMaxMarks > 0) ? (totalObtainedMarks / totalMaxMarks) * 100 : 0;
    }

    private String calculateGrade(double percentage) {
        if (Double.isNaN(percentage))
            return "N/A";
        if (percentage >= 90)
            return "A";
        if (percentage >= 80)
            return "B";
        if (percentage >= 70)
            return "C";
        if (percentage >= 60)
            return "D";
        return "F";
    }

    // New: Helper to get a cell's processed numerical value by its Excel-style
    // coordinates
    private double getCellValue(int rowCoord, int colCoord) {
        // Convert 1-based rowCoord to 0-based List index
        int dataRowIndex = rowCoord - 1;

        if (dataRowIndex < 0 || dataRowIndex >= data.size()) {
            return 0.0; // Row out of bounds
        }

        ObservableList<SimpleStringProperty> targetRow = data.get(dataRowIndex);

        // Column coordinates:
        // A=1 (ROLL NO, data index 1)
        // B=2 (NAME, data index 2)
        // C=3 (Subject 1, data index 3)
        // D=4 (Subject 2, data index 4) etc.
        // So, colCoord (1-based Excel column) directly maps to data index for Roll No,
        // Name, and subjects.
        // It's already adjusted correctly because ROLL NO is data index 1, NAME is 2.
        // And subjects start from data index 3 which aligns with 'C' (3rd letter).

        if (colCoord < 1 || colCoord >= targetRow.size()) { // Check against actual data properties available
            return 0.0; // Column out of bounds for the row
        }

        String cellValue = targetRow.get(colCoord).get();
        try {
            // Recursively process the cell value in case it's also a formula
            String processed = getProcessedValue(cellValue, dataRowIndex); // Pass the correct row index
            if (processed != null && !processed.isEmpty() && !processed.contains("Error")) {
                return Double.parseDouble(processed);
            }
        } catch (NumberFormatException e) {
            // Not a number, return 0.0 or throw an error for detailed reporting
        }
        return 0.0; // Default if not a valid number or error
    }

    // Modified: Processes cell input, supporting JavaScript expressions, cell
    // references, and SUM/AVERAGE functions
    private String getProcessedValue(String value, int currentRowIndex) {
        if (value == null || value.trim().isEmpty()) {
            return "";
        }

        if (value.startsWith("=")) {
            String formula = value.substring(1).trim();

            // 1. Handle SUM/AVERAGE functions (e.g., =SUM(C2:C4) or =AVERAGE(D1,D3))
            Matcher functionMatcher = FUNCTION_PATTERN.matcher(formula.toUpperCase());
            while (functionMatcher.find()) {
                String funcName = functionMatcher.group(1);
                String startColChar = functionMatcher.group(2);
                int startRowCoord = Integer.parseInt(functionMatcher.group(3));
                String endColChar = functionMatcher.group(4); // Null if not a range
                int endRowCoord = (functionMatcher.group(5) != null) ? Integer.parseInt(functionMatcher.group(5))
                        : startRowCoord;
                String additionalArgs = functionMatcher.group(6);

                List<Double> valuesToAggregate = new ArrayList<>();

                // Handle single cell or range
                Integer startColIdx = columnIndexMap.get(startColChar.charAt(0));
                Integer endColIdx = (endColChar != null) ? columnIndexMap.get(endColChar.charAt(0)) : startColIdx;

                if (startColIdx != null && endColIdx != null) {
                    int colStart = Math.min(startColIdx, endColIdx);
                    int colEnd = Math.max(startColIdx, endColIdx);
                    int rowStart = Math.min(startRowCoord, endRowCoord);
                    int rowEnd = Math.max(startRowCoord, endRowCoord);

                    for (int r = rowStart; r <= rowEnd; r++) {
                        for (int c = colStart; c <= colEnd; c++) {
                            valuesToAggregate.add(getCellValue(r, c));
                        }
                    }
                }

                // Handle additional comma-separated arguments (e.g., =SUM(C2,D3))
                if (additionalArgs != null && !additionalArgs.isEmpty()) {
                    String[] individualRefs = additionalArgs.substring(1).split(","); // remove leading comma and split
                    for (String ref : individualRefs) {
                        Matcher individualRefMatcher = CELL_REF_PATTERN.matcher(ref.trim().toUpperCase());
                        if (individualRefMatcher.matches()) {
                            String col = individualRefMatcher.group(1);
                            int row = Integer.parseInt(individualRefMatcher.group(2));
                            Integer colIdx = columnIndexMap.get(col.charAt(0)); // Only simple A-Z currently
                            if (colIdx != null) {
                                valuesToAggregate.add(getCellValue(row, colIdx));
                            }
                        }
                    }
                }

                double result = 0;
                if (!valuesToAggregate.isEmpty()) {
                    if (funcName.equals("SUM")) {
                        result = valuesToAggregate.stream().mapToDouble(Double::doubleValue).sum();
                    } else if (funcName.equals("AVERAGE")) {
                        result = valuesToAggregate.stream().mapToDouble(Double::doubleValue).average().orElse(0.0);
                    }
                }
                // Replace the function call in the formula string with its calculated result
                formula = functionMatcher.replaceFirst(String.format(Locale.US, "%.2f", result));
                functionMatcher = FUNCTION_PATTERN.matcher(formula.toUpperCase()); // Re-match for nested functions if
                                                                                   // any (simple case)
            }

            // 2. Handle simple cell references (e.g., =A1+B2)
            Matcher cellRefMatcher = CELL_REF_PATTERN.matcher(formula.toUpperCase());
            StringBuffer sb = new StringBuffer();
            while (cellRefMatcher.find()) {
                String col = cellRefMatcher.group(1);
                int row = Integer.parseInt(cellRefMatcher.group(2));

                Integer colIdx = columnIndexMap.get(col.charAt(0)); // Get internal index for 'A', 'B', 'C'...

                if (colIdx != null) {
                    double referencedValue = getCellValue(row, colIdx); // Fetch value
                    cellRefMatcher.appendReplacement(sb, String.format(Locale.US, "%.2f", referencedValue));
                } else {
                    cellRefMatcher.appendReplacement(sb, "0.0"); // Invalid column reference
                }
            }
            cellRefMatcher.appendTail(sb);
            formula = sb.toString();

            // 3. Evaluate the modified formula using ScriptEngine
            try {
                Object result = scriptEngine.eval(formula);
                if (result instanceof Number) {
                    return String.format(Locale.US, "%.2f", ((Number) result).doubleValue());
                }
                return result.toString();
            } catch (ScriptException e) {
                System.err.println("Script error for formula '" + value + "': " + e.getMessage());
                return "Error";
            } catch (Exception e) { // Catch other potential errors during parsing/conversion
                System.err.println("Formula processing error for '" + value + "': " + e.getMessage());
                return "Error";
            }
        }
        return value; // Not a formula, return as is
    }

    public class AlwaysEditingCell extends TableCell<ObservableList<SimpleStringProperty>, String> {
        private final TextField textField;

        public AlwaysEditingCell() {
            textField = new TextField();
            getStyleClass().add("editing-cell");

            textField.focusedProperty().addListener((obs, wasFocused, isFocused) -> {
                if (wasFocused && !isFocused) {
                    commitEdit(textField.getText());
                } else if (isFocused) {
                    getStyleClass().add("focused-cell"); // Add focused style
                } else {
                    getStyleClass().remove("focused-cell"); // Remove focused style
                }
            });
            textField.setOnAction(evt -> commitEdit(textField.getText()));

            textField.setOnKeyPressed(event -> {
                if (event.getCode() == null) {
                    return;
                }

                TableColumn<ObservableList<SimpleStringProperty>, ?> nextColumn = null;
                int currentRowIndex = getIndex();
                int newRowIndex = currentRowIndex;
                boolean handled = false;

                switch (event.getCode()) {
                    case UP:
                        newRowIndex--;
                        handled = true;
                        break;
                    case DOWN:
                        newRowIndex++;
                        handled = true;
                        break;
                    case LEFT:
                        nextColumn = getNextColumn(false);
                        handled = true;
                        break;
                    case RIGHT:
                        nextColumn = getNextColumn(true);
                        handled = true;
                        break;
                    case TAB:
                        nextColumn = getNextColumn(!event.isShiftDown());
                        handled = true;
                        break;
                    case ENTER:
                        commitEdit(textField.getText());
                        newRowIndex++;
                        handled = true;
                        break;
                    default:
                        break;
                }

                if (handled) {
                    event.consume();
                    if (nextColumn != null) {
                        getTableView().edit(currentRowIndex, nextColumn);
                    } else if (newRowIndex >= 0 && newRowIndex < getTableView().getItems().size()) {
                        getTableView().edit(newRowIndex, getTableColumn());
                    } else if (newRowIndex == getTableView().getItems().size()
                            && (event.getCode() == KeyCode.DOWN || event.getCode() == KeyCode.ENTER)) {
                        Platform.runLater(() -> {
                            addRowAt(data.size());
                            if (!data.isEmpty()) {
                                TableColumn<ObservableList<SimpleStringProperty>, ?> firstEditableCol = getFirstEditableColumn();
                                if (firstEditableCol != null) {
                                    getTableView().edit(data.size() - 1, firstEditableCol);
                                }
                            }
                        });
                    }
                }
            });
        }

        private TableColumn<ObservableList<SimpleStringProperty>, ?> getNextColumn(boolean forward) {
            List<TableColumn<ObservableList<SimpleStringProperty>, ?>> columns = getTableView().getVisibleLeafColumns();
            int currentIndex = columns.indexOf(getTableColumn());
            int nextIndex = currentIndex;
            while (true) {
                nextIndex = forward ? nextIndex + 1 : nextIndex - 1;
                if (nextIndex < 0 || nextIndex >= columns.size()) {
                    return null;
                }
                TableColumn<ObservableList<SimpleStringProperty>, ?> nextCol = columns.get(nextIndex);
                if (nextCol.isEditable() && (getTableView().getColumns().indexOf(nextCol) <= (2 + subjects.size()))) {
                    return nextCol;
                }
            }
        }

        private TableColumn<ObservableList<SimpleStringProperty>, ?> getFirstEditableColumn() {
            if (!getTableView().getColumns().isEmpty() && getTableView().getColumns().size() > 1) {
                TableColumn<ObservableList<SimpleStringProperty>, ?> rollNoCol = getTableView().getColumns().get(1);
                if (rollNoCol.isEditable()) {
                    return rollNoCol;
                }
            }
            return null;
        }

        @Override
        public void startEdit() {
            super.startEdit();
            if (isEmpty()) {
                return;
            }
            setGraphic(textField);
            // Display raw item value for editing, not the processed result
            textField.setText(getItem());
            Platform.runLater(() -> {
                textField.requestFocus();
                textField.selectAll();
            });
        }

        @Override
        public void cancelEdit() {
            super.cancelEdit();
            setText(getProcessedValue(getItem(), getIndex())); // Show processed value on cancel
            setGraphic(null);
            updateCellStyle(getItem()); // Apply style based on raw item value
        }

        @Override
        protected void updateItem(String item, boolean empty) {
            super.updateItem(item, empty);
            if (empty) {
                setText(null);
                setGraphic(null);
            } else {
                if (isEditing()) {
                    textField.setText(item);
                    setGraphic(textField);
                } else {
                    // Display the processed value when not editing
                    setText(getProcessedValue(item, getIndex()));
                    setGraphic(null);
                    updateCellStyle(item); // Apply style based on raw item value
                }
            }
        }

        // Applies "cell-error" style if the processed value indicates an error
        private void updateCellStyle(String item) {
            getStyleClass().remove("cell-error");
            // Check processed value for error indicator
            if (item != null && !item.isEmpty() && getProcessedValue(item, getIndex()).contains("Error")) {
                getStyleClass().add("cell-error");
            }
        }
    }

    public static class GradeCell extends TableCell<ObservableList<SimpleStringProperty>, String> {
        private final Label gradeLabel = new Label();

        public GradeCell() {
            gradeLabel.getStyleClass().add("grade-label");
        }

        @Override
        protected void updateItem(String item, boolean empty) {
            super.updateItem(item, empty);
            getStyleClass().removeAll("grade-a", "grade-b", "grade-c", "grade-d", "grade-f", "grade-n-a");
            if (item == null || empty) {
                setGraphic(null);
                setText(null);
            } else {
                gradeLabel.setText(item);
                getStyleClass().add("grade-" + item.toLowerCase().replace("/", ""));
                setGraphic(gradeLabel);
            }
        }
    }
}