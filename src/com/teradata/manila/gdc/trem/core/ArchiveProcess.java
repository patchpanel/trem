/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.teradata.manila.gdc.trem.core;

/**
 *
 * @author jl186034
 */
public class ArchiveProcess extends Thread {
    // --Commented out by Inspection (5/2/2016 8:33 PM):private static final Logger LOG = Logger.getLogger(ArchiveProcess.class.getName());

    private int rc;

    /**
     *
     */
    public ArchiveProcess() {
        rc = -1;
    }

// --Commented out by Inspection START (5/2/2016 8:33 PM):
//    /**
//     *
//     * @return
//     */
//    public int getRc() {
//        return rc;
//    }
// --Commented out by Inspection STOP (5/2/2016 8:33 PM)

    @Override
    public void run() {

        PropertiesFile pf = new PropertiesFile();

        ScriptRunner srArchiveOut = new ScriptRunner(pf, "powershell.exe -executionpolicy bypass -file ", BremConstants.ARCHIVE_OUT);
        ScriptRunner srArchiveIn = new ScriptRunner(pf, "powershell.exe -executionpolicy bypass -file ", BremConstants.ARCHIVE_IN);
        ScriptRunner srArchiveLog = new ScriptRunner(pf, "powershell.exe -executionpolicy bypass -file ", BremConstants.ARCHIVE_LOG);

        //Sequentially fire this processes since each is dependent on each other
        srArchiveOut.start();
        if (srArchiveOut.getRc() == 0) {
            srArchiveIn.start();
            if (srArchiveIn.getRc() == 0) {
                srArchiveLog.start();
                if (srArchiveLog.getRc() == 0) {
                    String now = new java.text.SimpleDateFormat("[MM/dd/yyyy HH:mm:ss]").format(new java.util.Date());
                    System.out.println(now + " Archiving complete.");
                    rc = 0;
                }
            }
        }
    }
}
