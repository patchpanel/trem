/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.teradata.manila.gdc.trem.util;

import com.teradata.manila.gdc.trem.core.PropertiesFile;
import com.teradata.manila.gdc.trem.core.ScriptRunner;
import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author jl186034
 */
public class Utilities {

    private static final Logger LOG = Logger.getLogger(Utilities.class.getName());

    /**
     *
     */
    public final static void killRunningProcess() {
        try {
            Process e = Runtime.getRuntime().exec("taskkill /F /IM excel.exe");
            //Process o = Runtime.getRuntime().exec("taskkill /F /IM outlook.exe");
            Process c = Runtime.getRuntime().exec("taskkill /F /IM cscript.exe");
            Process p = Runtime.getRuntime().exec("taskkill /F /IM powershell.exe");
        } catch (IOException ex) {
            java.util.logging.Logger.getLogger(ScriptRunner.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public final static void createPeriodLock() {
        String periodLockPath = getAppDataPath();

        PropertiesFile pf = new PropertiesFile();
        File fileGrpLock = new File(periodLockPath + "\\tremgrp." + pf.getLastBatchId().toUpperCase() + ".lck");
        File fileIndLock = new File(periodLockPath + "\\tremind." + pf.getLastBatchId().toUpperCase() + ".lck");

        try {
            fileGrpLock.createNewFile();
            fileIndLock.createNewFile();
        } catch (IOException ex) {
            Logger.getLogger(Utilities.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println(ex);
        }

    }

    public final static boolean checkPeriodClose() {
        PropertiesFile pf = new PropertiesFile();
        String periodLockPath = getAppDataPath();
        File fileGrpLock = new File(periodLockPath + "\\tremgrp." + pf.getLastBatchId().toUpperCase() + ".lck");
        File fileIndLock = new File(periodLockPath + "\\tremind." + pf.getLastBatchId().toUpperCase() + ".lck");

        return fileGrpLock.exists() && fileIndLock.exists();
    }

    private static String getAppDataPath() {
        String savePath = System.getenv("APPDATA") + "\\trem";

        File directory = new File(String.valueOf(savePath));
        if (!directory.exists()) {
            directory.mkdir();
        }
        return savePath;
    }
    
    private Utilities() {
    }
}
