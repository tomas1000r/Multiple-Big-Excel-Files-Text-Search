package sample;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.*;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.ImageView;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.BorderPane;
import javafx.stage.FileChooser;
import javafx.util.Callback;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.controlsfx.dialog.ProgressDialog;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;

public class Controller {

    @FXML
    BorderPane rootPane;

    @FXML
    ListView<String> fileLV;

    private ObservableList<String> filePaths;

    @FXML
    TextArea searchStringTA;

    @FXML
    TableView<ResultRow> resultTable;

    @FXML
    TableColumn<ResultRow, String> searchStringCol, atSheetCol, atFileCol;

    @FXML
    TableColumn<ResultRow, Integer> orderCol, atRowCol;

    @FXML
    CheckBox matchCaseCB;

    private ObservableList<ResultRow> resultRows;

    @FXML
    public void initialize()
    {
        filePaths = FXCollections.observableArrayList();
        fileLV.setItems(filePaths);
        resultRows = FXCollections.observableArrayList();
        prepareTable();
    }

    private void prepareTable()
    {
        resultTable.setItems(resultRows);
        resultTable.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);

        searchStringCol.setCellValueFactory(new PropertyValueFactory<ResultRow, String>("searchString"));
        atSheetCol.setCellValueFactory(new PropertyValueFactory<ResultRow, String>("sheetName"));
        atFileCol.setCellValueFactory(new PropertyValueFactory<ResultRow, String>("filePath"));
        atRowCol.setCellValueFactory(new PropertyValueFactory<ResultRow, Integer>("rowNumber"));
        orderCol.setCellFactory(new Callback<TableColumn<ResultRow, Integer>, TableCell<ResultRow, Integer>>() {

            @Override
            public TableCell<ResultRow, Integer> call(TableColumn<ResultRow, Integer> param) {
                return new TableCell<ResultRow, Integer>() {

                    @Override
                    protected void updateItem(Integer item, boolean empty) {
                        super.updateItem(item, empty);

                        Label label = new Label();
                        label.setAlignment(Pos.CENTER);
                        int index = getIndex() + 1;
                        label.setText(index + "");
                        setGraphic(label);
                    }
                };
            }
        });

    }

    public void selectXLSFiles()
    {
        FileChooser fileChooser = new FileChooser();

        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Excel files (*.xls, *.xlsx)", "*.xls", "*.xlsx");
        fileChooser.getExtensionFilters().add(extFilter);

        List<File> list = fileChooser.showOpenMultipleDialog(rootPane.getScene().getWindow());
        if (list == null) {
            return;
        }

        for (File f : list) {
            if (f != null) {
                if (!filePaths.contains(f.getAbsolutePath())) {
                    filePaths.add(f.getAbsolutePath());
                }
            }
        }
    }

    public void clearFileList()
    {
        fileLV.getItems().clear();
    }

    public void clearResults()
    {
        resultRows.clear();
    }

    public void copySelectionOnClick()
    {
        copySelectionToClipboard(resultTable);
    }

    /**
     * Start searching for matches in file
     * 1. Clear the table
     * 2. Add every match found
     */
    public void startSearching()
    {
        //clear the table
        resultRows.clear();
        final String[] stringToSearch = searchStringTA.getText().trim().split("\n");

        if (stringToSearch.length == 0) {
            System.out.println("no string to search");
            return;
        }

        final boolean matchCase = matchCaseCB.isSelected();

        Task<Void> searchingTask = new Task<Void>() {

            @Override
            protected Void call()
                throws Exception
            {
                for (String file : fileLV.getItems()) {
                    Workbook workbook;

                    try {
                        System.out.println("before making the book");
                        workbook = XcelFile.getBook(file);
                        System.out.println("working book get OK");
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        return null;
                    }

                    if (workbook == null) {
                        System.out.println("we are fucked!");
                        return null;
                    }

                    System.out.println("continue");

                    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                        Sheet sheet = workbook.getSheetAt(i);
                        Iterator<Row> iterator = sheet.iterator();

                        while (iterator.hasNext()) {

                            Row currentRow = iterator.next();
                            Iterator<Cell> cellIterator = currentRow.iterator();
                            StringBuilder rowContent = new StringBuilder();
                            String sheetName = sheet.getSheetName();
                            while (cellIterator.hasNext()) {
                                Cell currentCell = cellIterator.next();

                                //getCellTypeEnum shown as deprecated for version 3.15
                                //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                                if (currentCell.getCellTypeEnum() == CellType.STRING) {
                                    rowContent.append(currentCell.getStringCellValue());
                                }

                            }

                            for (String x : stringToSearch) {

                                if (!matchCase) {
                                    if (rowContent.toString().toLowerCase().contains(x.toLowerCase())) {
                                        resultRows.add(new ResultRow(x, file, sheetName, currentRow.getRowNum()));
                                        updateMessage("found " + x + " in " + file);
                                    }
                                } else {
                                    if (rowContent.toString().contains(x)) {
                                        resultRows.add(new ResultRow(x, file, sheetName, currentRow.getRowNum() + 1));
                                        updateMessage("found " + x + " in " + file);
                                    }
                                }
                            }
                        }
                    }
                }

                return null;
            }
        };

        ProgressDialog progressDialog = new ProgressDialog(searchingTask);
        progressDialog.setTitle("Searching...");

        new Thread(searchingTask).start();
    }

    @SuppressWarnings("rawtypes")
    public void copySelectionToClipboard(final TableView<?> table) {
        final Set<Integer> rows = new TreeSet<>();
        for (final TablePosition tablePosition : table.getSelectionModel().getSelectedCells()) {
            rows.add(tablePosition.getRow());
        }
        final StringBuilder strb = new StringBuilder();
        boolean firstRow = true;
        for (final Integer row : rows) {
            if (!firstRow) {
                strb.append('\n');
            }
            firstRow = false;
            boolean firstCol = true;
            for (final TableColumn<?, ?> column : table.getColumns()) {
                if (!firstCol) {
                    strb.append('\t');
                }
                firstCol = false;
                final Object cellData = column.getCellData(row);
                strb.append(cellData == null ? "" : cellData.toString());
            }
        }
        final ClipboardContent clipboardContent = new ClipboardContent();
        clipboardContent.putString(strb.toString());
        Clipboard.getSystemClipboard().setContent(clipboardContent);
    }
}
