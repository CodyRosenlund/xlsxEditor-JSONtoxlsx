//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by Fernflower decompiler)
//

package com.xlsxEditor.controller;

import com.xlsxEditor.view.editorUI;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.JFileChooser;

public class xlsxEditorController {
    private editorUI mainWindow;
    private JFileChooser fileChooser;
    private File inputFile;
    private volatile xlsxEditor xlsxEditor;
    private volatile Thread altThread;

    public xlsxEditorController(editorUI mainWindow) {
        this.mainWindow = mainWindow;
        this.mainWindow.addButtonActionListener(this.mainWindow.fileSelectButton, new xlsxEditorController.SelectInputFileListener());
        this.mainWindow.addButtonActionListener(this.mainWindow.runButton, new xlsxEditorController.RunButtonActionListener());
    }

    public void setInputFile(File inputFile) {
        this.inputFile = inputFile;
    }

    private void run() {
        this.xlsxEditor = new xlsxEditor(this.mainWindow.inputFileTextField.getText(), this.mainWindow.sheetNameTextField.getText(), this.mainWindow.columnTextField.getText(), this);
        this.altThread = new Thread(this.xlsxEditor);
        this.altThread.start();
    }

    public void updateFroRun(int exit_value) {
        this.altThread = null;
        if(exit_value == 0) {
            this.mainWindow.progressBar.setIndeterminate(false);
            this.mainWindow.statusLabel.setText("Completed");
        } else {
            this.mainWindow.progressBar.setIndeterminate(false);
            this.mainWindow.statusLabel.setText("Completed");
            String errorMessage = "An Error occurred.\n";
            if (exit_value == 1) {
                errorMessage += "Unable to find input file.";
            }
            else if (exit_value == 2) {
                errorMessage += "Unable to read input file.";
            }
            else if (exit_value == 3) {
                errorMessage += "Unable to to find worksheet.";
            } else if (exit_value == 4) {
                errorMessage += "Please close any related xlsx files.";
            } else if (exit_value == 5) {
                errorMessage += "Unable to to find worksheet.";
            } else {
                errorMessage += "Unexpected error.";
            }
            
            this.mainWindow.showAlert(errorMessage);
        }
    }

    private class ExitProgramListener implements ActionListener {
        private ExitProgramListener() {
        }

        public void actionPerformed(ActionEvent e) {
            System.exit(0);
        }
    }

    private class RunButtonActionListener implements ActionListener {
        private RunButtonActionListener() {
        }

        public void actionPerformed(ActionEvent e) {
            if(xlsxEditorController.this.inputFile == null && xlsxEditorController.this.mainWindow.inputFileTextField.getText().length() == 0) {
                xlsxEditorController.this.mainWindow.showAlert("Please select Input File.");
            } else if(xlsxEditorController.this.mainWindow.sheetNameTextField.getText().length() == 0) {
                xlsxEditorController.this.mainWindow.showAlert("Please enter tab or sheet name to be modified.");
            } else if(xlsxEditorController.this.mainWindow.columnTextField.getText().length() == 0) {
                xlsxEditorController.this.mainWindow.showAlert("Please enter desired column to be modified.");
            } else {
                xlsxEditorController.this.mainWindow.progressBar.setIndeterminate(true);
                xlsxEditorController.this.mainWindow.statusLabel.setText("Processing");
                xlsxEditorController.this.run();
            }
        }
    }

    private class SelectInputFileListener implements ActionListener {
        private SelectInputFileListener() {
        }

        public void actionPerformed(ActionEvent e) {
            xlsxEditorController.this.fileChooser = new JFileChooser();
            int returnVal = xlsxEditorController.this.fileChooser.showDialog(xlsxEditorController.this.mainWindow, "select");
            if(returnVal == 0) {
                String file = xlsxEditorController.this.fileChooser.getSelectedFile().getAbsolutePath();
                String path = xlsxEditorController.this.fileChooser.getSelectedFile().getParent();
                xlsxEditorController.this.mainWindow.inputFileTextField.setText(file);
                xlsxEditorController.this.setInputFile(xlsxEditorController.this.fileChooser.getSelectedFile());
            }
        }
    }

}
