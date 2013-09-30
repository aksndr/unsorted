package ru.aksndr.methods;

import com.crystaldecisions.report.web.viewer.CrystalReportViewer;
import com.crystaldecisions.sdk.occa.report.application.ParameterFieldController;
import com.crystaldecisions.sdk.occa.report.application.ReportClientDocument;
import com.crystaldecisions.sdk.occa.report.data.Fields;
import com.crystaldecisions.sdk.occa.report.data.IParameterField;
import com.crystaldecisions.sdk.occa.report.exportoptions.ReportExportFormat;
import com.crystaldecisions.sdk.occa.report.lib.ReportSDKException;
import com.crystaldecisions.sdk.occa.report.lib.ReportSDKExceptionBase;
import com.documentum.com.DfClientX;
import com.documentum.com.IDfClientX;
import com.documentum.drs.common.ReportDetailHandler;
import com.documentum.drs.common.ReportDetails;
import com.documentum.drs.method.internal.DocbaseController;
import com.documentum.drs.method.internal.trace.Tracer;
import com.documentum.fc.client.IDfClient;
import com.documentum.fc.client.IDfSession;
import com.documentum.fc.client.IDfSessionManager;
import com.documentum.fc.client.IDfSysObject;
import com.documentum.fc.common.*;
import com.documentum.fc.methodserver.IDfMethod;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.ResourceBundle;

/**
 * Author: a.arzamastsev Date: 01.11.12  Time: 15:57
 */
public class BuildContent implements IDfMethod {

    private static final String DQL_GET_TEMPLATE =
            "dm_sysobject where r_object_id = '%s'";
    private static final String R_FOLDER_PATH = "r_folder_path";
    private static final String ARGUMENTS_ERROR = "Parameter '%s' is required for method: 'mtd_build_content'";
    private static final String DOCBASE = "docbase";
    private static final String USER = "user";
    private static final String TICKET = "ticket";
    private static final String TICKET_FROM_SUPER_USER = "ticketFromSuperUser";
    private static final String DOCUMENT_ID = "documentId";
    private static final String REPORT_ID = "reportId";
    private static final String FORMAT = "format";
    private static final String RTF_EDITABLE_CONTENT_FORMAT = "rtf_editable";
    private static final String MS_WORD_CONTENT_FORMAT = "msw8";
    private static final String XLS_CONTENT_FORMAT = "excel8book";
    private static final String PDF_CONTENT_FORMAT = "pdf";
    private static final String DEFAULT_CONTENT_FORMAT = MS_WORD_CONTENT_FORMAT;
    private static final String PATH_SEPARATOR = "/";
    private static final int SESSION_ESTABLISHING_ATTEMPTS_COUNT = 10;
    private static final String[] SESSION_START_EXCEPTIONS = {"DM_SESSION_E_START_FAIL", "DM_SESSION_E_AUTH_FAIL"};

    private Map<String, String> m_params = new HashMap<String, String>();
    private DocbaseController m_docbaseController;

    @SuppressWarnings("RawUseOfParameterizedType")
    public int execute(Map parameters, PrintWriter printWriter) throws Exception {
        convertParams(parameters);
        initTicketParam();
        checkRequiredParams();

        initDocbaseController();

        try {
            runBuildContent(m_docbaseController.createDocbaseSession());
        } catch (DfException e) {
            throw new DfException("Failed to launch method runBuildContent. " + e);
        } finally {
            m_docbaseController.releaseDocbaseSession();
        }
        return 0;
    }

    private void initTicketParam() throws DfException {
        if (isGetTicketFromSuperuserSession()) {
            m_params.put(TICKET, getSuperUserSession().getLoginTicketForUser(m_params.get(USER)));
        }
    }

    private boolean isGetTicketFromSuperuserSession() {
        return m_params.containsKey(TICKET_FROM_SUPER_USER) && m_params.get(TICKET_FROM_SUPER_USER).equalsIgnoreCase("true");
    }

    private void initDocbaseController() {
        m_docbaseController = new DocbaseController(
                m_params.get(USER),
                m_params.get(TICKET),
                m_params.get(DOCBASE),
                null,
                new Tracer(System.out));
    }

    private void convertParams(Map params) {
        if (params == null) {
            throw new NullPointerException("Given params is null");
        }
        for (Object key : params.keySet()) {
            String[] values = (String[]) params.get(key);
            if (values == null || values.length == 0) {
                throw new IllegalArgumentException("Param value for " + key + " is empty");
            }
            m_params.put((String) key, values[0]);
        }
    }

    private void checkRequiredParams() {
        if (m_params == null || m_params.isEmpty()) {
            throw new NullPointerException("Method Arguments are null");
        }
        if (!m_params.containsKey(DOCBASE)) {
            throw new IllegalArgumentException(String.format(ARGUMENTS_ERROR, DOCBASE));
        }
        if (!m_params.containsKey(USER)) {
            throw new IllegalArgumentException(String.format(ARGUMENTS_ERROR, USER));
        }
        if (!m_params.containsKey(TICKET)) {
            throw new IllegalArgumentException(String.format(ARGUMENTS_ERROR, TICKET + "' or '" + TICKET_FROM_SUPER_USER));
        }
        if (!m_params.containsKey(DOCUMENT_ID)) {
            throw new IllegalArgumentException(String.format(ARGUMENTS_ERROR, DOCUMENT_ID));
        }
        if (!m_params.containsKey(REPORT_ID)) {
            throw new IllegalArgumentException(String.format(ARGUMENTS_ERROR, REPORT_ID));
        }
    }

    private void runBuildContent(IDfSession dfSession) throws DfException, IOException {

        ReportDetails reportDetails = createReportDetails(dfSession);
        ReportDetailHandler rdHandler = new ReportDetailHandler();
        ReportClientDocument reportClientDocument = rdHandler.updateConnectionDetails(reportDetails, true);

        CrystalReportViewer viewer = new CrystalReportViewer();
        try {
            viewer.setReportSource(reportClientDocument.getReportSource());
            reportClientDocument.getDatabaseController().logon(m_params.get(USER), dfSession.getLoginTicket());
            viewer.setEnableLogonPrompt(false);
        } catch (ReportSDKException e) {
            DfLogger.error(this, "Failed to set getDatabaseController ''{0}''",
                    new String[]{reportClientDocument.toString()}, e);
            throw new DfException("Failed to set getDatabaseController" + e.getMessage());
        } catch (ReportSDKExceptionBase reportSDKExceptionBase) {
            DfLogger.error(this, "Failed to set setReportSource ''{0}''",
                    new String[]{viewer.getName()}, reportSDKExceptionBase);
            throw new DfException("Failed to set setReportSource" + reportSDKExceptionBase);
        }
        try {
            Fields fields = reportClientDocument.getDataDefController().getDataDefinition().getParameterFields();
            ParameterFieldController paramFieldController =
                    reportClientDocument.getDataDefController().getParameterFieldController();
            if (paramFieldController != null) {
                if (fields != null) {

                    for (int x = 0; x < fields.size(); x++) {
                        try {
                            IParameterField parameterField = (IParameterField) fields.getField(x);
                            String paramName = parameterField.getName();
                            String paramValue = encodeParamValue(m_params.get(paramName));
                            paramFieldController.setCurrentValue("", paramName, paramValue);
                        } catch (ReportSDKException e) {
                            DfLogger.error(this, "Failed to set Param Field Value ''{0}''",
                                    new String[]{fields.getField(x).getName()}, e);
                            throw new DfException("Failed to set Param Field Value. " + e.getMessage());
                        }
                    }
                }
                ByteArrayInputStream byteArrayInputStream = null;
                ByteArrayOutputStream byteArrayOutputStream = null;
                try {
                    byteArrayInputStream = (ByteArrayInputStream)
                            reportClientDocument.getPrintOutputController().export(getExportFormat());

                    byte[] byteArray;
                    if (byteArrayInputStream != null) {
                        byteArray = new byte[byteArrayInputStream.available()];
                        byteArrayOutputStream = new ByteArrayOutputStream(byteArrayInputStream.available());
                        int x = byteArrayInputStream.read(byteArray, 0, byteArrayInputStream.available());
                        byteArrayOutputStream.write(byteArray, 0, x);
                        IDfSysObject obj = (IDfSysObject) dfSession.getObject(DfId.valueOf(m_params.get(DOCUMENT_ID)));
                        obj.setContentEx(byteArrayOutputStream, getFormat(), 0);
                        obj.save();
                    }
                } finally {
                    reportClientDocument.close();
                    if (byteArrayInputStream != null) {
                        byteArrayInputStream.close();
                    }
                    if (byteArrayOutputStream != null) {
                        byteArrayOutputStream.close();
                    }
                }
            }
        } catch (ReportSDKException e) {
            DfLogger.error(this, "Failed to get Fields from report ''{0}''", new String[]{""}, e);
            throw new DfException("Failed to get Fields from report " + e.getMessage());
        }
    }

    private String getFormat() {
        return m_params.get(FORMAT) == null || m_params.get(FORMAT).trim().length() == 0 ? DEFAULT_CONTENT_FORMAT : m_params.get(FORMAT);
    }

    private String exportToLocalFileSystem(IDfSession dfSession) throws DfException {
        return m_docbaseController.getDocumentFromRepository(getReportNameWithPath(dfSession));
    }

    private String getReportNameWithPath(IDfSession dfSession) throws DfException {
        String templateNameWithPath = null;
        String templateName;
        String templatePath;

        IDfSysObject templateObj = (IDfSysObject)
                dfSession.getObjectByQualification(String.format(DQL_GET_TEMPLATE, m_params.get(REPORT_ID)));
        if (templateObj == null) {
            throw new DfException("Failed to get object with id " + m_params.get(REPORT_ID));
        }
        templateName = templateObj.getObjectName();
        templatePath = (dfSession.getObject(templateObj.getFolderId(0))).getString(R_FOLDER_PATH) + PATH_SEPARATOR;
        if (templateName != null) {
            templateNameWithPath = templatePath + templateName;
        }
        return templateNameWithPath;
    }

    private ReportExportFormat getExportFormat() {
        final String format = getFormat();
        ReportExportFormat exportFormat = ReportExportFormat.MSWord;
        if (XLS_CONTENT_FORMAT.equalsIgnoreCase(format)) {
            exportFormat = ReportExportFormat.MSExcel;
        } else if (PDF_CONTENT_FORMAT.equalsIgnoreCase(format)) {
            exportFormat = ReportExportFormat.PDF;
        } else if (RTF_EDITABLE_CONTENT_FORMAT.equalsIgnoreCase(format)) {
            exportFormat = ReportExportFormat.editableRTF;
        }
        return exportFormat;
    }

    private ReportDetails createReportDetails(IDfSession dfSession) throws DfException {
        ReportDetails report = new ReportDetails();
        report.setDocbaseName(m_params.get(DOCBASE));
        report.setDomainName(null);
        report.setUserName(m_params.get(USER));
        report.setPassword(m_params.get(TICKET));
        report.setReportName(exportToLocalFileSystem(dfSession));
        return report;
    }

    private IDfSession getSuperUserSession() throws DfException {
        IDfClientX clientX = new DfClientX();
        IDfClient client = clientX.getLocalClient();
        IDfSessionManager sessionManager = client.newSessionManager();
        IDfLoginInfo loginInfo = clientX.getLoginInfo();
        ResourceBundle bundle = ResourceBundle.getBundle("SuperuserProp", Locale.ENGLISH);
        if (bundle == null) {
            throw new DfException("Failed to get resource bundle");
        }
        String login = bundle.getString("login");

        loginInfo.setUser(login);

        sessionManager.setIdentity(m_params.get(DOCBASE), loginInfo);

        IDfSession session = null;

        for (int i = 0; i < SESSION_ESTABLISHING_ATTEMPTS_COUNT; i++) {
            try {
                session = sessionManager.getSession(m_params.get(DOCBASE));
                break;
            } catch (DfException e) {
                if (i == SESSION_ESTABLISHING_ATTEMPTS_COUNT - 1 || !hasMessageIdsInChain(e, SESSION_START_EXCEPTIONS)) {
                    throw new DfException("Failed to establish connection with docbase " + m_params.get(DOCBASE), e);
                }
                sessionManager.clearIdentities();
                sessionManager.setIdentity(m_params.get(DOCBASE), loginInfo);
            }
        }
        return session;
    }

    private static boolean hasMessageIdsInChain(IDfException exception, String[] messageIds) {
        if (exception == null || messageIds == null || messageIds.length == 0) {
            return false;
        }

        //iterate throw exceptions chain
        IDfException ex = exception;
        do {
            for (String messageId : messageIds) {
                if (messageId.equals(ex.getMessageId())) {
                    return true;
                }
            }
            ex = ex.getNextException();
        } while (ex != null);

        return false;
    }

    private static String encodeParamValue(String value) {
        return value.replaceAll("&#x22", "\"").replaceAll("&#x27", "'");
    }
}