package sample;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.BorderPane;
import javafx.stage.FileChooser;
import javafx.util.Callback;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.controlsfx.dialog.ProgressDialog;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Controller {

    @FXML
    BorderPane rootPane;

    @FXML
    ListView<String> fileLV;

    ObservableList<String> filePaths;

    @FXML
    TextArea searchStringTA;


    @FXML
    TableView<ResultRow> resultTable;

    @FXML
    TableColumn<ResultRow, String> searchStringCol, atSheetCol, atFileCol;

    @FXML
    TableColumn<ResultRow, Integer> orderCol, atRowCol;


    private ObservableList<ResultRow> resultRows;

    @FXML
    CheckBox matchCaseCB;


    @FXML
    public void initialize()
    {
        filePaths = FXCollections.observableArrayList();
        fileLV.setItems(filePaths);
        resultRows = FXCollections.observableArrayList();
        prepareTable();
    }


    private static ArrayList<ResultRow> getMatch(Sheet sheet, String[] searchStrings, String filePath, boolean matchCase)
    {

        System.out.println("finding match");
        Iterator<Row> iterator = sheet.iterator();
        ArrayList<ResultRow> rows = new ArrayList<ResultRow>();

        System.out.println("match case is: " + matchCase);
        while (iterator.hasNext()) {

            Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            StringBuilder rowContent = new StringBuilder();
            String sheetName = sheet.getSheetName();
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();

                //getCellTypeEnum shown as deprecated for version 3.15
                //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                if (currentCell.getCellTypeEnum() == CellType.STRING)
                    rowContent.append(currentCell.getStringCellValue());

            }

            for (String x : searchStrings)
            {
                if (!matchCase)
                {
                    if (rowContent.toString().toLowerCase().contains(x.toLowerCase()))
                    {
                        rows.add(new ResultRow(
                                x, filePath, sheetName, currentRow.getRowNum()
                        ));
                        System.out.println(x + " available at: " + currentRow.getRowNum());
                    }
                } else
                {
                    if (rowContent.toString().contains(x))
                    {
                        rows.add(new ResultRow(
                                x, filePath, sheetName, currentRow.getRowNum()
                        ));
                        System.out.println(x + " available at: " + currentRow.getRowNum());
                    }
                }

            }



        }

        return rows;
    }

    private void prepareTable()
    {
        resultTable.setItems(resultRows);
        searchStringCol.setCellValueFactory(new PropertyValueFactory<ResultRow, String>("searchString"));
        atSheetCol.setCellValueFactory(new PropertyValueFactory<ResultRow, String>("sheetName"));
        atFileCol.setCellValueFactory(new PropertyValueFactory<ResultRow, String>("filePath"));
        atRowCol.setCellValueFactory(new PropertyValueFactory<ResultRow, Integer>("rowNumber"));
        orderCol.setCellFactory(new Callback<TableColumn<ResultRow, Integer>, TableCell<ResultRow, Integer>>() {
        @Override
        public TableCell<ResultRow, Integer> call(TableColumn<ResultRow, Integer> param) {
            return new TableCell<ResultRow, Integer>(){
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
        if (list == null)
            return;

        for (File f : list)
        {
            if (f!=null)
            {
                if (!filePaths.contains(f.getAbsolutePath()))
                    filePaths.add(f.getAbsolutePath());
            }
        }
    }

    public void clearFileList()
    {
        fileLV.getItems().clear();
    }

    public void startSearching()
    {
        final String[] stringToSearch = searchStringTA.getText().trim().split("\n");

        if (stringToSearch.length == 0)
        {
            System.out.println("no string to search");
            return;
        }


        final boolean matchCase = matchCaseCB.isSelected();

        Task<Void> searchingTask = new Task<Void>() {
            @Override
            protected Void call() throws Exception {
                for (String file : fileLV.getItems())
                {
                    Workbook workbook;

                    try
                    {
                        System.out.println("before making the book");
                        workbook =  XcelFile.getBook(file);
                        System.out.println("working book get OK");
                    } catch (Exception ex)
                    {
                        ex.printStackTrace();
                        return null;
                    }

                    if (workbook == null)
                    {
                        System.out.println("we are fucked!");
                        return null;
                    }

                    System.out.println("continue");

                    for (int i = 0; i < workbook.getNumberOfSheets(); i++)
                    {
                        resultRows.addAll(getMatch(workbook.getSheetAt(i), stringToSearch, file, matchCase));
                    }

                }

                return null;
            }
        };

        ProgressDialog progressDialog = new ProgressDialog(searchingTask);

        progressDialog.setTitle("Searching...");

        new Thread(searchingTask).start();



    }
}
