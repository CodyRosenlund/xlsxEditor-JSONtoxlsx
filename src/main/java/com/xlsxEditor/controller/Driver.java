package com.xlsxEditor.controller;

import com.xlsxEditor.view.editorUI;

/**
 * Hello world!
 */

public class Driver {

    public static void main(String[] args) {
        editorUI editorUI = new editorUI();
        new xlsxEditorController(editorUI);
    }
}