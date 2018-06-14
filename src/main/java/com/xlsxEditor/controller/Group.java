package com.xlsxEditor.controller;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Cody Rosenlund, SPS Commerce
 * June 2018
 * <p>
 * Group.java
 * <p>
 * A simple class to track a group's xpath and its min/max which is just converted to a mandatory field if it doesnt
 * start with zero (0/1000 is not mandatory)
 */
class Group {

    private String xpath;
    private String mandatory;

    Group(String xpath, String mandatory) {
        this.xpath = xpath;
        this.mandatory = mandatory != null && !mandatory.equals("") && mandatory.trim().matches("^[1-9].*") ? "m" : "";
    }


    String getXpath() {
        return xpath;
    }

    String getMandatory() {
        return mandatory;
    }

    /**
     * Returns a list of the group's xpath split by the '/'
     */
    List<String> getXpathAsList() {
        return new ArrayList<>(Arrays.asList(xpath.trim().replaceAll("^/", "").split("/")));
    }
}
