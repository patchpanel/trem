/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.teradata.manila.gdc.trem.core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Paths;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author jl186034
 */
public class PropertiesFile {
    private static final Logger LOG = Logger.getLogger(PropertiesFile.class.getName());

    private String _rootDir;
    private String _inputDir;
    private String _outputDir;
    private String _tempDir;
    private String _logDir;
    private String _binDir;
    private String _excelBadgeReport;
    private String _excelResourceList;
    private String _txtResourcelist;
    private String _txtManagerList;
    private String _summarySheet;
    private String _detailEntExPairSheet;
    private String _detailRawSheet;
    private String _tagMngrRept;
    private String _lastBatchId;
    private String _extractResourceListScript;
    private String _extractIndividualScript;
    private String _extractManagerScript;
    private String _emailAllScript;
    private String _emailIndividualScript;
    private String _emailManagerScript;
    private String _emailFrom;
    private String _smtpServer;
    private String _emailBody;
    private String _emailTo;
    private String _archiveDays;
    private String _archiveScript;
    private String _archiveDir;

    /**
     *
     */
    public PropertiesFile() {
        this.ReadPropertiesFile();
    }

    /**
     *
     * @return
     */
    public String getRootDir() {
        return _rootDir;
    }

    /**
     *
     * @param _rootDir
     */
    public void setRootDir(String _rootDir) {
        this._rootDir = _rootDir;
    }

    /**
     *
     * @return
     */
    public String getInputDir() {
        return _inputDir;
    }

    /**
     *
     * @param _inputDir
     */
    public void setInputDir(String _inputDir) {
        this._inputDir = _inputDir;
    }

    /**
     *
     * @return
     */
    public String getOutputDir() {
        return _outputDir;
    }

    /**
     *
     * @param _outputDir
     */
    public void setOutputDir(String _outputDir) {
        this._outputDir = _outputDir;
    }

    /**
     *
     * @return
     */
    public String getTempDir() {
        return _tempDir;
    }

    /**
     *
     * @param _tempDir
     */
    public void setTempDir(String _tempDir) {
        this._tempDir = _tempDir;
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
    public String getBinDir() {
        return _binDir;
    }

    /**
     *
     * @param _binDir
     */
    public void setBinDir(String _binDir) {
        this._binDir = _binDir;
    }

    /**
     *
     * @return
     */
    public String getExcelBadgeReport() {
        return _excelBadgeReport;
    }

    /**
     *
     * @param _excelBadgeReport
     */
    public void setExcelBadgeReport(String _excelBadgeReport) {
        this._excelBadgeReport = _excelBadgeReport;
    }

    /**
     *
     * @return
     */
    public String getExcelResourceList() {
        return _excelResourceList;
    }

    /**
     *
     * @param _excelResourceList
     */
    public void setExcelResourceList(String _excelResourceList) {
        this._excelResourceList = _excelResourceList;
    }

    /**
     *
     * @return
     */
    public String getTxtResourcelist() {
        return _txtResourcelist;
    }

    /**
     *
     * @param _txtResourcelist
     */
    public void setTxtResourcelist(String _txtResourcelist) {
        this._txtResourcelist = _txtResourcelist;
    }

    /**
     *
     * @return
     */
    public String getTxtManagerList() {
        return _txtManagerList;
    }

    /**
     *
     * @param _txtManagerList
     */
    public void setTxtManagerList(String _txtManagerList) {
        this._txtManagerList = _txtManagerList;
    }

    /**
     *
     * @return
     */
    public String getSummarySheet() {
        return _summarySheet;
    }

    /**
     *
     * @param _summarySheet
     */
    public void setSummarySheet(String _summarySheet) {
        this._summarySheet = _summarySheet;
    }

    /**
     *
     * @return
     */
    public String getDetailEntExPairSheet() {
        return _detailEntExPairSheet;
    }

    /**
     *
     * @param _detailEntExPairSheet
     */
    public void setDetailEntExPairSheet(String _detailEntExPairSheet) {
        this._detailEntExPairSheet = _detailEntExPairSheet;
    }

    /**
     *
     * @return
     */
    public String getDetailRawSheet() {
        return _detailRawSheet;
    }

    /**
     *
     * @param _detailRawSheet
     */
    public void setDetailRawSheet(String _detailRawSheet) {
        this._detailRawSheet = _detailRawSheet;
    }

    /**
     *
     * @return
     */
    public String getTagMngrRept() {
        return _tagMngrRept;
    }

    /**
     *
     * @param _tagMngrRept
     */
    public void setTagMngrRept(String _tagMngrRept) {
        this._tagMngrRept = _tagMngrRept;
    }

    /**
     *
     * @return
     */
    public String getLastBatchId() {
        return _lastBatchId;
    }

    /**
     *
     * @param _lastBatchId
     */
    public void setLastBatchId(String _lastBatchId) {
        this._lastBatchId = _lastBatchId;
    }

    /**
     *
     * @return
     */
    public String getExtractResourceListScript() {
        return _extractResourceListScript;
    }

    /**
     *
     * @param _extractResourceListScript
     */
    public void setExtractResourceListScript(String _extractResourceListScript) {
        this._extractResourceListScript = _extractResourceListScript;
    }

    /**
     *
     * @return
     */
    public String getExtractIndividualScript() {
        return _extractIndividualScript;
    }

    /**
     *
     * @param _extractIndividualScript
     */
    public void setExtractIndividualScript(String _extractIndividualScript) {
        this._extractIndividualScript = _extractIndividualScript;
    }

    /**
     *
     * @return
     */
    public String getExtractManagerScript() {
        return _extractManagerScript;
    }

    /**
     *
     * @param _extractManagerScript
     */
    public void setExtractManagerScript(String _extractManagerScript) {
        this._extractManagerScript = _extractManagerScript;
    }

    /**
     *
     * @return
     */
    public String getEmailAllScript() {
        return _emailAllScript;
    }

    /**
     *
     * @param _emailAllScript
     */
    public void setEmailAllScript(String _emailAllScript) {
        this._emailAllScript = _emailAllScript;
    }

    /**
     *
     * @return
     */
    public String getEmailIndividualScript() {
        return _emailIndividualScript;
    }

    /**
     *
     * @param _emailIndividualScript
     */
    public void setEmailIndividualScript(String _emailIndividualScript) {
        this._emailIndividualScript = _emailIndividualScript;
    }

    /**
     *
     * @return
     */
    public String getEmailManagerScript() {
        return _emailManagerScript;
    }

    /**
     *
     * @param _emailManagerScript
     */
    public void setEmailManagerScript(String _emailManagerScript) {
        this._emailManagerScript = _emailManagerScript;
    }

    /**
     *
     * @return
     */
    public String getEmailFrom() {
        return _emailFrom;
    }

    /**
     *
     * @param _emailFrom
     */
    public void setEmailFrom(String _emailFrom) {
        this._emailFrom = _emailFrom;
    }

    /**
     *
     * @return
     */
    public String getSmtpServer() {
        return _smtpServer;
    }

    /**
     *
     * @param _smtpServer
     */
    public void setSmtpServer(String _smtpServer) {
        this._smtpServer = _smtpServer;
    }

    /**
     *
     * @return
     */
    public String getEmailBody() {
        return _emailBody;
    }

    /**
     *
     * @param _emailBody
     */
    public void setEmailBody(String _emailBody) {
        this._emailBody = _emailBody;
    }

    /**
     *
     * @return
     */
    public String getEmailTo() {
        return _emailTo;
    }

    /**
     *
     * @param _emailTo
     */
    public void setEmailTo(String _emailTo) {
        this._emailTo = _emailTo;
    }

    /**
     *
     * @return
     */
    public String getArchiveDays() {
        return _archiveDays;
    }

    /**
     *
     * @param _archiveDays
     */
    public void setArchiveDays(String _archiveDays) {
        this._archiveDays = _archiveDays;
    }

    /**
     *
     * @return
     */
    public String getArchiveScript() {
        return _archiveScript;
    }

    /**
     *
     * @param _archiveScript
     */
    public void setArchiveScript(String _archiveScript) {
        this._archiveScript = _archiveScript;
    }

    /**
     *
     * @return
     */
    public String getArchiveDir() {
        return _archiveDir;
    }

    /**
     *
     * @param _archiveDir
     */
    public void setArchiveDir(String _archiveDir) {
        this._archiveDir = _archiveDir;
    }

    /**
     *
     */
    public void ReadPropertiesFile() {
        Properties props = new Properties();

        InputStream is = null;
        try {
            File file = new File(Paths.get(".").toAbsolutePath().normalize().toString() + "/config.xml");
            is = new FileInputStream(file);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(PropertiesFile.class.getName()).log(Level.SEVERE, null, ex);
        }
        try {
            //load the xml file into properties format
            props.loadFromXML(is);
        } catch (IOException ex) {
            Logger.getLogger(PropertiesFile.class.getName()).log(Level.SEVERE, null, ex);
        }
        _rootDir = props.getProperty("rootdir");
        _inputDir = props.getProperty("inputdir");
        _outputDir = props.getProperty("outputdir");
        _tempDir = props.getProperty("tempdir");
        _logDir = props.getProperty("logdir");
        _binDir = props.getProperty("bindir");
        _excelBadgeReport = props.getProperty("excelbadgereport");
        _excelResourceList = props.getProperty("excelresourcelist");
        _txtResourcelist = props.getProperty("txtresourcelist");
        _txtManagerList = props.getProperty("txtmanagerlist");
        _summarySheet = props.getProperty("summarysheet");
        _detailEntExPairSheet = props.getProperty("detailentexpairsheet");
        _detailRawSheet = props.getProperty("detailrawsheet");
        _tagMngrRept = props.getProperty("tagmngrrept");
        _lastBatchId = props.getProperty("lastbatchid");
        _extractResourceListScript = props.getProperty("extractresourcelistscript");
        _extractIndividualScript = props.getProperty("extractindividualscript");
        _extractManagerScript = props.getProperty("extractmanagerscript");
        _emailAllScript = props.getProperty("emailallscript");
        _emailIndividualScript = props.getProperty("emailindividualscript");
        _emailManagerScript = props.getProperty("emailmanagerscript");
        _emailFrom = props.getProperty("emailfrom");
        _smtpServer = props.getProperty("smtpserver");
        _emailBody = props.getProperty("emailbody");
        _emailTo = props.getProperty("emailto");
        _archiveDays = props.getProperty("archivedays");
        _archiveScript = props.getProperty("archivescript");
        _archiveDir = props.getProperty("archivedir");
    }

    /**
     *
     */
    public void WritePropertiesFile() {
        Properties props = new Properties();

        props.setProperty("rootdir", _rootDir);
        props.setProperty("inputdir", _inputDir);
        props.setProperty("outputdir", _outputDir);
        props.setProperty("tempdir", _tempDir);
        props.setProperty("logdir", _logDir);
        props.setProperty("bindir", _binDir);
        props.setProperty("excelbadgereport", _excelBadgeReport);
        props.setProperty("excelresourcelist", _excelResourceList);
        props.setProperty("txtresourcelist", _txtResourcelist);
        props.setProperty("txtmanagerlist", _txtManagerList);
        props.setProperty("summarysheet", _summarySheet);
        props.setProperty("detailentexpairsheet", _detailEntExPairSheet);
        props.setProperty("detailrawsheet", _detailRawSheet);
        props.setProperty("tagmngrrept", _tagMngrRept);
        props.setProperty("lastbatchid", _lastBatchId);
        props.setProperty("extractresourcelistscript", _extractResourceListScript);
        props.setProperty("extractindividualscript", _extractIndividualScript);
        props.setProperty("extractmanagerscript", _extractManagerScript);
        props.setProperty("emailallscript", _emailAllScript);
        props.setProperty("emailindividualscript", _emailIndividualScript);
        props.setProperty("emailmanagerscript", _emailManagerScript);
        props.setProperty("emailfrom", _emailFrom);
        props.setProperty("smtpserver", _smtpServer);
        props.setProperty("emailbody", _emailBody);
        props.setProperty("emailto", _emailTo);
        props.setProperty("archivedays", _archiveDays);
        props.setProperty("archivescript", _archiveScript);
        props.setProperty("archivedir", _archiveDir);

        OutputStream os = null;
        try {
            File file = new File(Paths.get(".").toAbsolutePath().normalize().toString() + "/config.xml");
            os = new FileOutputStream(file);
            System.out.println("Property Saved");
        } catch (FileNotFoundException ex) {
            Logger.getLogger(PropertiesFile.class.getName()).log(Level.SEVERE, null, ex);
        }

        try {
            //store the properties detail into a pre-defined XML file
            props.storeToXML(os, "Directories", "UTF-8");

        } catch (IOException ex) {
            Logger.getLogger(PropertiesFile.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
