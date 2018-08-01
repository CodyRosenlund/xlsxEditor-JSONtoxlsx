package com.xlsxEditor.controller;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;

import org.json.JSONArray;
import org.json.JSONObject;
import org.json.JSONTokener;

/**
 * Cody Rosenlund, SPS Commerce
 * June 2018
 * <p>
 * xlsxEditor.java
 * <p>
 * Does the brunt of the work. Parses through the xlsx file and creates the retailers
 */
public class xlsxEditor implements Runnable {

    private String path;
    private String output_path = "";
    private String rsxVersion;
    private String docType;
    private xlsxEditorController xdc;
    private LinkedList<Node> linkedNodes;

    public xlsxEditor(String path, xlsxEditorController xdc) {
        linkedNodes = new LinkedList<>();
        this.path = path;
        this.xdc = xdc;
        String[] path_split = path.split("\\.");
        for (int s = 0; s < path_split.length - 2; s++) {
            output_path += path_split[s];
        }
    }

    public void run() {
        int exit_value = 0;

        try {

            File myFile = new File(path);
            FileInputStream fis = new FileInputStream(myFile);
            fis.close();
        } catch (FileNotFoundException e) {
            exit_value = 1;
        } catch (IOException e) {
            exit_value = 2;
        }

        JSONObject json = null;
        try {
            File myFile = new File(path);
            FileInputStream fis = new FileInputStream(myFile);
            JSONTokener jtokener = new JSONTokener(fis);
            json = new JSONObject(jtokener);
            fis.close();
        } catch (NullPointerException e) {
            exit_value = 5;
        } catch (FileNotFoundException e) {
            exit_value = 1;
        } catch (IOException e) {
            exit_value = 2;
        }

        //get root info first to start the tree object
        if (json != null && json.has("name")) {
            Node<JSONObject> rootNode = new Node(json.getString("name"), json.getString("name"), "", "", 1, 1, "1", "", "", "", null, "", true, null);

            linkedNodes.add(rootNode);
//            rootNode.version =  designMeta.viewedSchema.version
            parseJSON(json.getJSONArray("children"), rootNode);
        }

        if (!linkedNodes.isEmpty()) {
            new xlsxBuilder().buildxlsx(linkedNodes);
        }

        //return any errors after xpaths are (attempted) added
        this.xdc.updateFroRun(exit_value);
    }

    private void parseJSON(JSONArray json, Node parentNode) {
        for (Object jobj : json) {

            String name = ((JSONObject) jobj).getString("name");
            String xpath = parentNode.xpath + "/" + name;
            String definition = "";
            String dataType = "";
            int minLength = 1;
            int maxLength = 1;
            String maxOccurs = "1";
            String usage = "";
            HashMap<String, Qualifier> qualifiers = new HashMap<>();
            String notes = "";
            String mandatoryFor = "";
            String optionalFor = "";
            boolean visible = false;
            if (((JSONObject) jobj).has("consolidatedDocumentation")) {
                definition = ((JSONObject) jobj).getString("consolidatedDocumentation");
            }
            if (((JSONObject) jobj).has("attributes")) {
                for (Object attributes : ((JSONObject) jobj).getJSONArray("attributes")) {
                    if (((JSONObject) attributes).has("displayName")) {
                        dataType = ((JSONObject) attributes).getString("displayName");
                    }
                    if (((JSONObject) attributes).has("minLength")) {
                        minLength = ((JSONObject) attributes).getInt("minLength");
                    }
                    if (((JSONObject) attributes).has("maxLength")) {
                        maxLength = ((JSONObject) attributes).getInt("maxLength");
                    }

                }
            }
            if (((JSONObject) jobj).has("visible")) {
                visible = ((JSONObject) jobj).getBoolean("visible");
            }
            if (((JSONObject) jobj).has("minOccurs")) {
                usage = ((JSONObject) jobj).getInt("minOccurs") == 1 ? "M" : "O";
            }
            if (((JSONObject) jobj).has("consolidatedMandatoryFor")) {
                mandatoryFor = ((JSONObject) jobj).getString("consolidatedMandatoryFor");
            }
            if (((JSONObject) jobj).has("consolidatedOptionalFor")) {
                optionalFor = ((JSONObject) jobj).getString("consolidatedOptionalFor");
            }
            if (((JSONObject) jobj).has("maxOccurs")) {
                maxOccurs = ((JSONObject) jobj).getString("maxOccurs");
            }
            if (((JSONObject) jobj).has("consolidatedQualifierDescriptions")) {
                for (Object quals : ((JSONObject) jobj).getJSONArray("consolidatedQualifierDescriptions")) {
                    String value = "";
                    String description = "";
                    String qualMandatoryFor = "";
                    String qualOptionalFor = "";
                    if (((JSONObject) quals).has("enum")) {
                        value = ((JSONObject) quals).getString("enum");
                    }
                    if (((JSONObject) quals).has("documentation")) {
                        description = ((JSONObject) quals).getString("documentation");
                    }
                    if (((JSONObject) quals).has("mandatoryFor")) {
                        qualMandatoryFor = ((JSONObject) quals).getString("mandatoryFor");
                    }
                    if (((JSONObject) quals).has("optionalFor")) {
                        qualOptionalFor = ((JSONObject) quals).getString("optionalFor");
                    }
                    qualifiers.put(value, new Qualifier(value, description, qualMandatoryFor, qualOptionalFor));
                }
            }
            if (((JSONObject) jobj).has("notes")) {
                notes = ((JSONObject) jobj).getString("notes");
            }

            Node<JSONObject> newChild = new Node(xpath, name, definition, dataType, minLength, maxLength, maxOccurs, usage, mandatoryFor, optionalFor, qualifiers, notes, visible, parentNode);

            parentNode.addChild(newChild);

            //add to this linkedList so we can add to the xlsx in a certain order
            linkedNodes.add(newChild);

            if (((JSONObject) jobj).has("children")) {
                parseJSON(((JSONObject) jobj).getJSONArray("children"), newChild);
            }
        }
    }
}