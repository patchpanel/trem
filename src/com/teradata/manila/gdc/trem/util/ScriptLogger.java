/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.teradata.manila.gdc.trem.util;

import java.util.logging.Logger;
import javax.swing.JTextArea;

/**
 *
 * @author jl186034
 */
public class ScriptLogger {
    private static final Logger LOG = Logger.getLogger(ScriptLogger.class.getName());

    private String _logDir;
    private javax.swing.JTextArea textArea;

    /**
     *
     * @param logFile
     * @param textArea
     */
    public ScriptLogger(String logFile, JTextArea textArea) {
        this._logDir = logFile;
        this.textArea = textArea;
        saveLog();
    }

    /**
     *
     * @return
     */
    public String getLogDir() {
        return _logDir;
    }

    /**
     *
     * @param _logDir
     */
    public void setLogDir(String _logDir) {
        this._logDir = _logDir;
    }

    /**
     *
     * @return
     */
    public JTextArea getTextArea() {
        return textArea;
    }

    /**
     *
     * @param textArea
     */
    public void setTextArea(JTextArea textArea) {
        this.textArea = textArea;
    }

    private void saveLog() {
        java.io.BufferedWriter outFile = null;
        try {
            outFile = new java.io.BufferedWriter(new java.io.FileWriter(_logDir));
            textArea.write(outFile);

        } catch (java.io.IOException ex) {
        } finally {
            if (outFile != null) {
                try {
                    outFile.close();
                } catch (java.io.IOException e) {
                }
            }
        }
    }
}
