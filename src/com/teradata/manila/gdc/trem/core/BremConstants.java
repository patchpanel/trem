/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.teradata.manila.gdc.trem.core;

import java.util.logging.Logger;

/**
 *
 * @author jl186034
 */
public class BremConstants {

    /**
     *
     */
    public static final int EXTRACT_RESOURCE_LIST = 1;

    /**
     *
     */
    public static final int EXTRACT_INDIVIDUAL_LIST = 2;

    /**
     *
     */
    public static final int EXTRACT_MANAGER_LIST = 3;

    /**
     *
     */
    public static final int EMAIL_SINGLE = 4;

    /**
     *
     */
    public static final int EMAIL_GROUP = 5;

    /**
     *
     */
    public static final int EMAIL_ALL = 6;

    /**
     *
     */
    public static final int ARCHIVE_OUT = 7;

    /**
     *
     */
    public static final int ARCHIVE_IN = 8;

    /**
     *
     */
    public static final int ARCHIVE_LOG = 9;
    //public static final boolean PROCESS_ALL = true;
    private static final Logger LOG = Logger.getLogger(BremConstants.class.getName());
    private BremConstants() {
    }
}
