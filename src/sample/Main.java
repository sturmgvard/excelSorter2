package sample;

import javafx.application.Application;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.text.*;
import javafx.scene.text.Font;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Main extends Application {
    private static HashMap<Workbook, String> wbMap = new HashMap<>();
    private static HashMap<Row, String> rowsMap = new HashMap<>();
    static Button buttonAddFile;
    static StringBuilder sbAddedFiles = new StringBuilder("Добавленные файлы:\n");

    static Text txtAddedFiles = new Text(sbAddedFiles.toString());

    static TextField inputText;
    static Button btnStartSearch;
    static GridPane gridPaneOutput;

    static {
        inputText = new TextField();
        inputText.setPrefColumnCount(10);

        gridPaneOutput = new GridPane();
//        gridPaneOutput.setGridLinesVisible(true);
        gridPaneOutput.setHgap(20);
        gridPaneOutput.setVgap(20);


    }

    @Override
    public void start(Stage primaryStage) throws Exception {
        FileChooser fileChooser = new FileChooser();
        BorderPane root = new BorderPane();
        ScrollPane scrollPane = new ScrollPane();


        scrollPane.setContent(root);

        scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
        scrollPane.setHbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);

        VBox leftVBox = new VBox();
        leftVBox.setBorder(new Border(new BorderStroke(Color.DEEPSKYBLUE,//Color.valueOf("#00C6D5")
                BorderStrokeStyle.SOLID, new CornerRadii(5), new BorderWidths(3))));
        root.setLeft(leftVBox);
        leftVBox.setSpacing(20);
        leftVBox.setPadding(new Insets(10, 10, 10, 10));


        VBox rightVBox = new VBox();
        rightVBox.setBorder(new Border(new BorderStroke(Color.DEEPSKYBLUE,
                BorderStrokeStyle.SOLID, new CornerRadii(5), new BorderWidths(3))));
        root.setRight(rightVBox);
        rightVBox.setSpacing(15);
        rightVBox.setPadding(new Insets(10, 10, 10, 10));


        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("xlsx Files", "*.xlsx")
        );
        buttonAddFile = new Button("Добавить excel файл");
        buttonAddFile.setMinWidth(110);
        buttonAddFile.setOnAction(e -> {
            File selectedFile = fileChooser.showOpenDialog(primaryStage);
            if (selectedFile != null) {
                try (FileInputStream fis = new FileInputStream(selectedFile)) {
                    Workbook wb = new XSSFWorkbook(fis);


                    wbMap.put(wb, selectedFile.getName());
                    sbAddedFiles.append("\n" + selectedFile.getName());
                    txtAddedFiles.setText(sbAddedFiles.toString());
                } catch (IOException exc) {

                }
            }
        });
        leftVBox.getChildren().addAll(buttonAddFile, txtAddedFiles);


        btnStartSearch = new Button("Поиск по этому слову");
        btnStartSearch.addEventHandler(MouseEvent.MOUSE_CLICKED, new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent mouseEvent) {
                int rowIndex = 0;
                int columnIndex = 0;
                if (inputText != null) {
                    fillRowsList(inputText.getText());


                    for (Map.Entry<Row, String> entry : rowsMap.entrySet()) {
                        for (Cell cell : entry.getKey()) {
                            if (cell != null) {
                                if (cell.getCellType() != CellType.BLANK) {
                                    if (cell.getCellType() == CellType.NUMERIC) {
                                        Text t = new Text("цена: " + cell.toString());
                                        t.setFont(new Font(18));
                                        t.setFill(Color.RED);

                                        gridPaneOutput.add(t, columnIndex++, rowIndex);//todo add format string

                                    } else {
                                        Text t = new Text();
                                        t.setText(cell.toString());
                                        t.setFont(new Font(18));

                                        gridPaneOutput.add(t, columnIndex++, rowIndex);//todo add format string

                                    }
                                }


                            }
                        }
                        Text t = new Text();
                        t.setText(entry.getValue() + " , номер строки = " + (entry.getKey().getRowNum() + 1));
                        t.setFill(Color.BLUE);
                        gridPaneOutput.add(t, columnIndex, rowIndex);
                        columnIndex = 0;
                        ++rowIndex;

                    }

                }
            }
        });
        rightVBox.getChildren().addAll(inputText, btnStartSearch, gridPaneOutput);


        primaryStage.setTitle("Эксель поиск (@sedof)");
        primaryStage.setScene(new Scene(scrollPane, 1700, 477));
        primaryStage.show();
    }


    public static void main(String[] args) {
        launch(args);
    }


    public static void fillRowsList(String search) {
        if (!wbMap.isEmpty()) {
            for (Map.Entry<Workbook, String> entry : wbMap.entrySet()) {
                for (Sheet sheet : entry.getKey()) {
                    for (Row row : sheet) {
                        for (Cell cell : row) {
                            if (cell != null) {
                                if (cell.getCellType() == CellType.STRING) {
                                    if (cell.toString().toLowerCase().contains(search.toLowerCase())) {
//                                        String s = cell.getStringCellValue().toUpperCase() + " имя файла: "+ entry.getValue() + " номер строки: " + (row.getRowNum()+1);
//                                        cell.setCellValue(s);
                                        rowsMap.put(row, entry.getValue());

                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }


//echo "# excelSorter2" >> README.md
//        git init
//        git add README.md
//        git commit -m "first commit"
//        git remote add origin https://github.com/sturmgvard/excelSorter2.git
//        git push -u origin master
}
