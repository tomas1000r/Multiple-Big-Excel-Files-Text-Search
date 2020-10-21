package sample;

import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;

public class ResultRow {
    SimpleStringProperty searchString = new SimpleStringProperty(), filePath = new SimpleStringProperty(), sheetName = new SimpleStringProperty();
    SimpleIntegerProperty rowNumber = new SimpleIntegerProperty();

    public ResultRow(String _searchString, String _filePath, String _sheetName, int _rowNumber)
    {
        setFilePath(_filePath);
        setSearchString(_searchString);
        setRowNumber(_rowNumber);
        setSheetName(_sheetName);
    }

    public String getSheetName() {
        return sheetName.get();
    }

    public SimpleStringProperty sheetNameProperty() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName.set(sheetName);
    }

    public String getSearchString() {
        return searchString.get();
    }

    public SimpleStringProperty searchStringProperty() {
        return searchString;
    }

    public void setSearchString(String searchString) {
        this.searchString.set(searchString);
    }

    public String getFilePath() {
        return filePath.get();
    }

    public SimpleStringProperty filePathProperty() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath.set(filePath);
    }

    public int getRowNumber() {
        return rowNumber.get();
    }

    public SimpleIntegerProperty rowNumberProperty() {
        return rowNumber;
    }

    public void setRowNumber(int rowNumber) {
        this.rowNumber.set(rowNumber);
    }
}
