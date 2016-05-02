package com.teradata.manila.gdc.trem.core;

import java.util.logging.Logger;

/**
 *
 * @author jl186034
 */
public class EmailProcess extends Thread {
    private static final Logger LOG = Logger.getLogger(EmailProcess.class.getName());

    private final javax.swing.JButton button;

    /**
     *
     * @param button
     */
    public EmailProcess(final javax.swing.JButton button) {
        this.button = button;
    }

    @Override
    public void run() {

        PropertiesFile pf = new PropertiesFile();

        ScriptRunner srEmailAll = new ScriptRunner(pf, "powershell.exe -executionpolicy bypass -file ", BremConstants.EMAIL_ALL);
        ScriptRunner srEmailSingle = new ScriptRunner(pf, "powershell.exe -executionpolicy bypass -file ", BremConstants.EMAIL_SINGLE);

        //Sequentially fire this processes since each is dependent on each other
        button.setEnabled(false);
        srEmailAll.start();
        if (srEmailAll.getRc() == 0) {
            srEmailSingle.start();
            String now = new java.text.SimpleDateFormat("[MM/dd/yyyy HH:mm:ss]").format(new java.util.Date());
            if (srEmailSingle.getRc() == 0) {
                System.out.println(now + " Email process Complete");
            }
        }
        button.setEnabled(true);
    }
}
