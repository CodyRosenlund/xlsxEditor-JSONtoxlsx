package com.xlsxEditor.controller;

import java.util.ArrayList;
import java.util.List;

/**
 * Cody Rosenlund, SPS Commerce
 * June 2018
 * <p>
 * Retailer.java
 * <p>
 * A simple class to track groups and fields
 */
class Retailer {

    private List<Field> fieldsList;
    private List<Group> groupList;
    private String retailerName;
    private String retailerHubID;

    Retailer(String retailerName, String retailerHubID) {
        this.retailerName = retailerName;
        this.retailerHubID = retailerHubID;
        fieldsList = new ArrayList<>();
        groupList = new ArrayList<>();
    }

    void addField(String xpath, String mandatory, String qualifiers, String retailerInfo) {
        Field field = new Field(xpath, mandatory, qualifiers, retailerInfo);
        fieldsList.add(field);
    }

    void addGroup(String xpath, String retailerInfo) {
        Group group = new Group(xpath, retailerInfo);
        groupList.add(group);
    }

    List<Field> getRetailerFields() {
        return fieldsList;
    }

    List<Group> getRetailerGroups() {
        return groupList;
    }

    String getRetailerName() {
        return retailerName;
    }

    String getRetailerHubID() {
        return retailerHubID;
    }


}
