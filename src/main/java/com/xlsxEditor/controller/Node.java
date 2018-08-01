package com.xlsxEditor.controller;

import java.util.HashMap;
import java.util.LinkedList;
import java.util.Map;

class Node<T> {
    String xpath;
    String name;
    String definition;
    String dataType;
    int minLength;
    int maxLength;
    String maxOccurs;
    String usage;
    String mandatoryFor;
    String optionalFor;
    HashMap<String, Qualifier> qualifiers;
    String designNotes;
    private Node<T> parent;
    boolean visible;
    private LinkedList<Node<T>> children;
    int hierarchyOffSet;

    Node(String xpath, String name, String definition, String dataType, int minLength, int maxLength, String maxOccurs, String usage, String mandatoryFor, String optionalFor, HashMap<String, Qualifier> qualifiers, String designNotes, boolean visible, Node<T> parent) {
        this.xpath = xpath;
        this.name = name;
        this.definition = definition;
        this.dataType = dataType;
        this.minLength = minLength;
        this.maxLength = maxLength;
        this.maxOccurs = maxOccurs;
        this.usage = usage;
        this.mandatoryFor = mandatoryFor;
        this.optionalFor = optionalFor;
        this.qualifiers = qualifiers == null ? new HashMap<>() : qualifiers;
        this.designNotes = designNotes;
        this.parent = parent;
        this.visible = visible;
        this.children = new LinkedList<>();
        hierarchyOffSet = (parent != null ? parent.hierarchyOffSet + 3 : 0);
    }

    void addChild(Node<T> child) {
        children.add(child);
    }

    LinkedList<Node<T>> getChildren() {
        return children;
    }

    boolean isField() {
        return children.isEmpty();
    }

    boolean firstChildIsAGroup() {
        if (!children.isEmpty()) {
            if (children.get(0).isField()) {
                return true;
            }
//            for (Node child: children) {
//                if (!child.isField()) {
//                    return true;
//                }
//            }
        }
        return false;
    }

    String getQualifiersAndDescriptions() {
        StringBuilder qualString = new StringBuilder();
        for (Map.Entry<String, Qualifier> qualPair : qualifiers.entrySet()) {
            qualString.append(qualPair.getKey());
            qualString.append(": ");
            qualString.append(qualPair.getValue().description);
            qualString.append("\r\n");
        }
        return qualString.toString();
    }
}