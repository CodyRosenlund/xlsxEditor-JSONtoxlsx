package com.xlsxEditor.controller;

import org.json.*;

import java.io.*;
import java.util.*;

/**
 * Cody Rosenlund, SPS Commerce
 * June 2018
 * <p>
 * JSONBuilder.java
 * <p>
 * Our JSON 'editor'. Reads in a json file, traverses it to find groups and fields and adjusts them according to
 * their object attributes.
 */
class JSONBuilder {
    private JSONObject retailerJSON = null;
    private Retailer currentRetailer;

    int createFiles(String rsxVersion, String docType, HashMap<String, Retailer> retailersHash, String outputPath) {

        for (Map.Entry entry : retailersHash.entrySet()) {
            currentRetailer = (Retailer) entry.getValue();
            InputStream jsonFile;
            try {
                jsonFile = getClass().getClassLoader().getResourceAsStream(rsxVersion.replaceAll("\\.", "_") + "_" + docType.toLowerCase() + ".json");
                JSONTokener jtokener = new JSONTokener(jsonFile);
                retailerJSON = new JSONObject(jtokener);
            } catch (NullPointerException e) {
                System.out.println(e);
                return 5;
            }

            // for each field, search in the json and updated if found
            for (Field field : currentRetailer.getRetailerFields()) {

                retailerJSON = updateFieldInJSON(retailerJSON, field, field.getXpathAsList());
            }

            // for each group, search in the json and updated if found
            for (Group group : currentRetailer.getRetailerGroups()) {

                retailerJSON = updateGroupInJSON(retailerJSON, group, group.getXpathAsList());
            }

            // create the actual new JSON with a generated name
            createJSONFile(outputPath, rsxVersion, docType);

        }
        return 0;
    }

    private JSONObject updateFieldInJSON(JSONObject jsonObject, Field field, List<String> fieldXpath) {

        List<String> updatedFieldXpath = new ArrayList(fieldXpath);
        updatedFieldXpath.remove(0);
        for (String path : fieldXpath) {
            if (jsonObject.getString("name").equals(path) && jsonObject.has("children")) {
                return jsonObject.put("children", findFieldInJSON(jsonObject.getJSONArray("children"), field, updatedFieldXpath));
            }
        }

        return jsonObject;
    }

    //recursive method that modifies Json
    private JSONArray findFieldInJSON(JSONArray jarray, Field field, List<String> fieldXpath) {
        //step through array and find the next path, if it has children
        for (String path : fieldXpath) {
            for (Object jobj : jarray) {
                //if we have found the field we want to change
                if (((JSONObject) jobj).get("name").equals(path)) {

                    ((JSONObject) jobj).put("visible", true);
                    if (fieldXpath.indexOf(path) == (fieldXpath.size() - 1)) {
                        //change the json here
                        ((JSONObject) jobj).put("visible", field.getMandatory().length() > 0);
                        ((JSONObject) jobj).put("minOccurs", field.getMandatory().equalsIgnoreCase("m") ? "1" : "0");
                        ((JSONObject) jobj).put("qualifiers", field.getQualifers());
                        ((JSONObject) jobj).put("notes", field.getRetailerInfo());

                        return jarray;
                    }
                    //else we shorten the xpath list and get a smaller json to try again
                    else if (((JSONObject) jobj).has("children")) {
                        List<String> updatedFieldXpath = new ArrayList(fieldXpath);
                        updatedFieldXpath.remove(0);

                        ((JSONObject) jobj).put("children", findFieldInJSON(((JSONObject) jobj).getJSONArray("children"), field, updatedFieldXpath));
                        return jarray;
                    }
                }
            }
        }

        return jarray;
    }

    private JSONObject updateGroupInJSON(JSONObject jsonObject, Group group, List<String> groupXpath) {

        List<String> updatedFieldXpath = new ArrayList(groupXpath);
        updatedFieldXpath.remove(0);
        for (String path : groupXpath) {
            if (jsonObject.getString("name").equals(path) && jsonObject.has("children")) {
                return jsonObject.put("children", findGroupInJSON(jsonObject.getJSONArray("children"), group, updatedFieldXpath));
            }
        }

        return jsonObject;
    }

    //recursive method that modifies Json
    private JSONArray findGroupInJSON(JSONArray jarray, Group group, List<String> groupXpath) {
        //step through array and find the next path, if it has children
        for (String path : groupXpath) {
            for (Object jobj : jarray) {
                //if we have found the group we want to change
                if (((JSONObject) jobj).get("name").equals(path)) {

                    if (groupXpath.indexOf(path) == (groupXpath.size() - 1)) {
                        //change the json here
                        ((JSONObject) jobj).put("minOccurs", group.getMandatory().equalsIgnoreCase("m") ? "1" : "0");

                        return jarray;
                    }
                    //else we shorten the xpath list and get a smaller json to try again
                    else if (((JSONObject) jobj).has("children")) {
                        List<String> updatedGroupXpath = new ArrayList(groupXpath);
                        updatedGroupXpath.remove(0);

                        ((JSONObject) jobj).put("children", findGroupInJSON(((JSONObject) jobj).getJSONArray("children"), group, updatedGroupXpath));
                        return jarray;
                    }
                }
            }
        }

        return jarray;
    }


    private void createJSONFile(String outputPath, String rsxVersion, String docType) {
        String string = JSONWriter.valueToString(retailerJSON);
        try {
            FileWriter file = new FileWriter(outputPath + currentRetailer.getRetailerName().replaceAll("[^a-zA-Z0-9]", "") + "_" + currentRetailer.getRetailerHubID().replaceAll("\\.0+", "") + "_" + rsxVersion.replaceAll("\\.", "_") + "_" + docType.toLowerCase() + ".json");
            file.write(string);
            file.close();
        } catch (Exception e) {
            System.err.println("Error: " + e);
        }
    }
}
