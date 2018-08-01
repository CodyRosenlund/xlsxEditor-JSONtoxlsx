package com.xlsxEditor.controller;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

public class Qualifier {

    String value;
    String description;
    Set<String> mandatoryFor;
    Set<String> optionalFor;

    Qualifier(String value, String description, String mandatoryFor, String optionalFor) {
        this.value = value;
        this.description = description;
        this.mandatoryFor = new HashSet<>(Arrays.asList(splitPartners(mandatoryFor)));
        this.optionalFor = new HashSet<>(Arrays.asList(splitPartners(optionalFor)));
    }

    private String[] splitPartners(String partnerString) {
        return partnerString.trim().split(",");
    }
}
