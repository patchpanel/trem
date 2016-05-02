package com.teradata.manila.gdc.trem.core;

import java.util.logging.Logger;

/**
 *
 * @author jl186034
 */
public class WholeProcess extends Thread {

    private static final Logger LOG = Logger.getLogger(WholeProcess.class.getName());

    private final javax.swing.JButton button;

    /**
     *
     * @param button
     */
    public WholeProcess(final javax.swing.JButton button) {
        this.button = button;
    }

    @Override
    public void run() {

        PropertiesFile pf = new PropertiesFile();

        ScriptRunner srExtractResourceList = new ScriptRunner(pf, "cscript.exe", BremConstants.EXTRACT_RESOURCE_LIST);
        ScriptRunner srExtractIndividualReport = new ScriptRunner(pf, "cscript.exe", BremConstants.EXTRACT_INDIVIDUAL_LIST);
        ScriptRunner srExtractGroupReport = new ScriptRunner(pf, "cscript.exe", BremConstants.EXTRACT_MANAGER_LIST);
        ScriptRunner srEmailAll = new ScriptRunner(pf, "powershell.exe -executionpolicy bypass -file ", BremConstants.EMAIL_ALL);
        ScriptRunner srEmailSingle = new ScriptRunner(pf, "powershell.exe -executionpolicy bypass -file ", BremConstants.EMAIL_SINGLE);
        //Sequentially fire this processes since each is dependent on each other
        button.setEnabled(false);
        srExtractResourceList.start();
        if (srExtractResourceList.getRc() == 0) {
            srExtractIndividualReport.start();
        }
        if (srExtractIndividualReport.getRc() == 0
                && srExtractResourceList.getRc() == 0) {
            srExtractGroupReport.start();
        }
        if (srExtractGroupReport.getRc() == 0
                && srExtractIndividualReport.getRc() == 0
                && srExtractResourceList.getRc() == 0) {
            srEmailAll.start();
        }
        if (srExtractGroupReport.getRc() == 0
                && srExtractIndividualReport.getRc() == 0
                && srExtractResourceList.getRc() == 0
                && srEmailAll.getRc() == 0) {
            srEmailSingle.start();
            String now = new java.text.SimpleDateFormat("[MM/dd/yyyy HH:mm:ss]").format(new java.util.Date());
            if (srEmailSingle.getRc() == 0) {
                System.out.println(now + " All processes complete");
            }
        }
        button.setEnabled(true);
    }
}
