package com.teradata.manila.gdc.trem.core;

/**
 *
 * @author jl186034
 */
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.logging.Logger;
import javax.swing.JTextArea;

/**
 *
 * @author jl186034
 */
public class ScriptRunner {
    private static final Logger LOG = Logger.getLogger(ScriptRunner.class.getName());

    private String _args;
    private javax.swing.JTextArea _textArea;
    private PropertiesFile _propertiesFile;
    private String _command;
    private int _option;
    private int _rc;

    /**
     *
     * @param propertiesFile
     * @param command
     * @param option
     */
    public ScriptRunner(PropertiesFile propertiesFile, String command, int option) {
        this._propertiesFile = propertiesFile;
        this._command = command;
        this._option = option;
    }

    /**
     *
     * @return
     */
    public int getRc() {
        return _rc;
    }

    private void setRc(int _rc) {
        this._rc = _rc;
    }

    /**
     *
     * @return
     */
    public PropertiesFile getAlProperties() {
        return _propertiesFile;
    }

    /**
     *
     * @param _alProperties
     */
    public void setAlProperties(PropertiesFile _alProperties) {
        this._propertiesFile = _alProperties;
    }

    /**
     *
     * @return
     */
    public int getOption() {
        return _option;
    }

    /**
     *
     * @param _option
     */
    public void setOption(int _option) {
        this._option = _option;
    }

    /**
     *
     * @return
     */
    public String getArgs() {
        return _args;
    }

    /**
     *
     * @param _args
     */
    public void setArgs(String _args) {
        this._args = _args;
    }

    /**
     *
     * @return
     */
    public JTextArea getTextArea() {
        return _textArea;
    }

    /**
     *
     * @param _textArea
     */
    public void setTextArea(JTextArea _textArea) {
        this._textArea = _textArea;
    }

    /**
     *
     * @return
     */
    public String getCommand() {
        return _command;
    }

    /**
     *
     * @param _command
     */
    public void setCommand(String _command) {
        this._command = _command;
    }

    /**
     *
     */
    public void start() {
        this._rc = this.runCommand();
    }

    private int runCommand() {
        int rc = -1;
        try {
            String line;

            OutputStream stdin = null;
            InputStream stderr = null;
            InputStream stdout = null;

            //Create the argument string
            this._args = this.buildCommand();
            System.out.println("Executing: " + this._args);
            Process process = Runtime.getRuntime().exec(this._args);
            stdin = process.getOutputStream();
            stderr = process.getErrorStream();
            stdout = process.getInputStream();

            // clean up if any output in stdout
            BufferedReader brCleanUp = new BufferedReader(new InputStreamReader(stdout));
            while ((line = brCleanUp.readLine()) != null) {
                System.out.println(line);
            }

            brCleanUp.close();

            // clean up if any output in stderr
            brCleanUp = new BufferedReader(new InputStreamReader(stderr));
            while ((line = brCleanUp.readLine()) != null) {
                System.out.println(line);
            }
            brCleanUp.close();
            System.out.println("RC=" + process.waitFor());

            stdin.close();
            stdout.close();
            stderr.close();
            rc = process.exitValue();
        } catch (IOException | InterruptedException e) {
            System.out.println(e);
        }
        return rc;
    }

    /**
     *
     * @return
     */
    public String buildCommand() {
        String cmd = null;

        String rootDir = _propertiesFile.getRootDir();
        String inputDir = _propertiesFile.getInputDir();
        String outputDir = _propertiesFile.getOutputDir();
        String tempDir = _propertiesFile.getTempDir();
        String logDir = _propertiesFile.getLogDir();
        String binDir = _propertiesFile.getBinDir();
        String excelBadgeReport = _propertiesFile.getExcelBadgeReport();
        String excelResourceList = _propertiesFile.getExcelResourceList();
        String txtResourcelist = _propertiesFile.getTxtResourcelist();
        String txtManagerList = _propertiesFile.getTxtManagerList();
        String summarySheet = _propertiesFile.getSummarySheet();
        String detailEntExPairSheet = _propertiesFile.getDetailEntExPairSheet();
        String detailRawSheet = _propertiesFile.getDetailRawSheet();
        String tagMngrRept = _propertiesFile.getTagMngrRept();
        String lastBatchId = _propertiesFile.getLastBatchId();
        String extractResourceListScript = _propertiesFile.getExtractResourceListScript();
        String extractIndividualScript = _propertiesFile.getExtractIndividualScript();
        String extractManagerScript = _propertiesFile.getExtractManagerScript();
        String emailAllScript = _propertiesFile.getEmailAllScript();
        String emailIndividualScript = _propertiesFile.getEmailIndividualScript();
        String emailManagerScript = _propertiesFile.getEmailManagerScript();
        String emailFrom = _propertiesFile.getEmailFrom();
        String smtpServer = _propertiesFile.getSmtpServer();
        String emailBody = _propertiesFile.getEmailBody();
        String emailTo = _propertiesFile.getEmailTo();
        String archiveDays = _propertiesFile.getArchiveDays();
        String archiveScript = _propertiesFile.getArchiveScript();
        String archiveDir = _propertiesFile.getArchiveDir();

        switch (this._option) {
            case BremConstants.EXTRACT_RESOURCE_LIST:
                //cmd = new String[]{"\"" + _command.toLowerCase() + "\"," + " \"" + _propertiesFile.get(15) + "\" \"" + _propertiesFile.get(7) + "\""};
                cmd = "\"" + _command.toLowerCase() + "\""
                        + " \"" + binDir + "/" + extractResourceListScript + "\""
                        + " \"" + inputDir + "/" + excelResourceList + "\""
                        + " \"" + inputDir + "/" + txtResourcelist + "\""
                        + " \"" + inputDir + "/" + txtManagerList + "\""
                        + " \"" + lastBatchId + "\"";
                break;
            case BremConstants.EXTRACT_INDIVIDUAL_LIST:
                //cscript ExtractIndividualReports.vbs "C:\atri\in\GDC Manila Resource List template v1 0.txt" "C:\atri\in\201603 - BDG_TimeReport_V2.xlsx" "Summary" "Detailed Entry Exit Pair" "Detailed Raw" "C:\atri\out" 201604
                cmd = "\"" + _command.toLowerCase() + "\""
                        + " \"" + binDir + "/" + extractIndividualScript + "\""
                        + " \"" + inputDir + "/" + txtResourcelist + "\""
                        + " \"" + inputDir + "/" + excelBadgeReport + "\""
                        + " \"" + summarySheet + "\""
                        + " \"" + detailEntExPairSheet + "\""
                        + " \"" + detailRawSheet + "\""
                        + " \"" + outputDir + "\""
                        + " \"" + lastBatchId + "\"";
                break;
            case BremConstants.EXTRACT_MANAGER_LIST:
                //cscript ExtractGroupReports.vbs "C:\atri\in\GDC Manila Resource List template v1 0.txt" "C:\atri\in\201603 - BDG_TimeReport_V2.xlsx" "Summary" "Detailed Entry Exit Pair" "Detailed Raw" "C:\atri\out" 201604 Practice
                cmd = "\"" + _command.toLowerCase() + "\""
                        + " \"" + binDir + "/" + extractManagerScript + "\""
                        + " \"" + inputDir + "/" + txtResourcelist + "\""
                        + " \"" + inputDir + "/" + txtManagerList + "\""
                        + " \"" + inputDir + "/" + excelBadgeReport + "\""
                        + " \"" + summarySheet + "\""
                        + " \"" + detailEntExPairSheet + "\""
                        + " \"" + detailRawSheet + "\""
                        + " \"" + outputDir + "\""
                        + " \"" + lastBatchId + "\""
                        + " \"" + tagMngrRept + "\"";
                break;
            case BremConstants.EMAIL_ALL:
                //powershell -ExecutionPolicy ByPass -File .\mailer.ps1 "C:\atri\in\GDC Manila Resource List template v1 0.txt" "C:\atri\in\GDC Manila Resource List template v1 0.Managers.txt" "C:\atri\out" jl186034@teradata.com outlook.td.teradata.com 201604 "Practice" "THIS IS A TEST"
                cmd = _command.toLowerCase()
                        + " " + binDir + "/" + emailAllScript
                        + " \"" + inputDir + "/" + txtResourcelist + "\""
                        + " \"" + inputDir + "/" + txtManagerList + "\""
                        + " \"" + outputDir + "\""
                        + " \"" + emailFrom + "\""
                        + " \"" + smtpServer + "\""
                        + " \"" + lastBatchId + "\""
                        + " \"" + tagMngrRept + "\""
                        + " \"" + emailBody + "\""
                        + " \"" + logDir + "\"";
                break;
//            case BremConstants.EMAIL_GROUP:
//                //powershell -ExecutionPolicy ByPass -File .\mailer.ps1 "C:\atri\in\GDC Manila Resource List template v1 0.txt" "C:\atri\in\GDC Manila Resource List template v1 0.Managers.txt" "C:\atri\out" jl186034@teradata.com outlook.td.teradata.com 201604 "Practice" "THIS IS A TEST"
//                cmd = new String[]{"\"" + _command.toLowerCase() + " \"," + " \"" + _propertiesFile.getValueAt(18, 2) + "\""
//                    + " \"" + _propertiesFile.getValueAt(8, 2) + "\""
//                    + " \"" + _propertiesFile.getValueAt(9, 2) + "\""
//                    + " \"" + _propertiesFile.getValueAt(6, 2) + "\""
//                    + " \"" + _propertiesFile.getValueAt(2, 2) + "\""
//                    + " \"" + _propertiesFile.getValueAt(22, 2) + "\""
//                    + " \"" + _propertiesFile.getValueAt(23, 2) + "\""
//                    + " \"" + _propertiesFile.getValueAt(14, 2) + "\""
//                    + " \"" + _propertiesFile.getValueAt(13, 2) + "\""
//                    + " \"" + _propertiesFile.getValueAt(24, 2) + "\""};
//                break;
            case BremConstants.EMAIL_SINGLE:
                //powershell -ExecutionPolicy ByPass -File .\SendMailIndividual.ps1 "C:\atri\in" jl186034@teradata.com jl186034@teradata.com localhost 201604 "THIS A TEST"
                cmd = _command.toLowerCase()
                        + " " + binDir + "/" + emailIndividualScript
                        + " \"" + inputDir + "/" + excelBadgeReport + "\""
                        + " \"" + emailFrom + "\""
                        + " \"" + emailTo + "\""
                        + " \"" + smtpServer + "\""
                        + " \"" + lastBatchId + "\""
                        + " \"" + emailBody + "\""
                        + " \"" + logDir + "\""
                        + " \"" + tagMngrRept + "\"";
                break;
            case BremConstants.ARCHIVE_OUT:
                //powershell -ExecutionPolicy ByPass -File .\archiver.ps1 "c:\brem\out" "c:\brem\archive" "c:\brem\log" "OUT" 1
                //powershell -ExecutionPolicy ByPass -File .\archiver.ps1 "c:\brem\in" "c:\brem\archive" "c:\brem\log" "IN" 1
                //powershell -ExecutionPolicy ByPass -File .\archiver.ps1 "c:\brem\log" "c:\brem\archive" "c:\brem\log" "LOG" 1
                cmd = _command.toLowerCase()
                        + " " + binDir + "/" + archiveScript
                        + " \"" + outputDir + "\""
                        + " \"" + archiveDir + "\""
                        + " \"" + logDir + "\""
                        + " \"OUT\""
                        + " " + archiveDays;
                break;
            case BremConstants.ARCHIVE_IN:
                cmd = _command.toLowerCase()
                        + " " + binDir + "/" + archiveScript
                        + " \"" + inputDir + "\""
                        + " \"" + archiveDir + "\""
                        + " \"" + logDir + "\""
                        + " \"IN\""
                        + " " + archiveDays;
                break;
            case BremConstants.ARCHIVE_LOG:
                cmd = _command.toLowerCase()
                        + " " + binDir + "/" + archiveScript
                        + " \"" + logDir + "\""
                        + " \"" + archiveDir + "\""
                        + " \"" + logDir + "\""
                        + " \"LOG\""
                        + " " + archiveDays;
                break;
        }
        return cmd;
    }
}
