package com.xlsxEditor.controller;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Cody Rosenlund, SPS Commerce
 * June 2018
 * <p>
 * Field.java
 * <p>
 * A simple class to track a field's xpath, m/o/c, qualifiers (which get cleaned up), and other notes
 */
class Field {

    private String xpath;
    private String mandatory;
    private String qualifers;
    private String retailerInfo;

    Field(String xpath, String mandatory, String qualifiers, String retailerInfo) {
        this.xpath = xpath;
        this.mandatory = mandatory;
        this.qualifers = cleanupQualifiers(qualifiers);
        this.retailerInfo = retailerInfo;
    }


    String getXpath() {
        return xpath;
    }

    String getMandatory() {
        return mandatory;
    }

    String getQualifers() {
        return qualifers;
    }

    String getRetailerInfo() {
        return retailerInfo;
    }

    /**
     * Cleans up qualifiers. Most qualifiers are sent something like: -ST- Shipto address
     * Where we only want the ST to be part of the qualifier string (comma separated)
     */
    private String cleanupQualifiers(String qualifiers) {
        String cleanedQuals = "";
        //remove all none space and alpha numeric characters
        qualifiers = qualifiers.replaceAll("[^a-zA-Z0-9\\s]", " ");
        String[] quals = qualifiers.split("\\n");
        for (String qual : quals) {
            cleanedQuals = (cleanedQuals.equals("") ? qual.trim().split("\\s", 2)[0] : cleanedQuals + "," + qual.trim().split("\\s", 2)[0].replaceAll("[^a-zA-Z0-9]", ""));
        }

        return (cleanedQuals.matches("") ? qualifiers : cleanedQuals);
    }

    /**
     * Returns a list of the field's xpath split by the '/'
     */
    List<String> getXpathAsList() {
        return new ArrayList<>(Arrays.asList(xpath.trim().replaceAll("^/", "").split("/")));
    }
}
