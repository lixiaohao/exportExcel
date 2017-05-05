package com.zhongzhou.web.excel;

import javax.validation.ConstraintViolation;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

/**
 * Created by lixiaohao on 2017/5/5
 *
 * @Description
 * @Create 2017-05-05 16:52
 * @Company
 */
public class JsonResponse {
    private static final long serialVersionUID = 571211199485773026L;
    private boolean success = true;
    private boolean doRedirect = false;
    private String redirectUrl;
    private String actionMessage;
    private String actionType;
    private Object data;
    private long totalResult;
    private Map<String, Object> otherAttributes;

    public JsonResponse() {
    }

    public <T> JsonResponse(Set<ConstraintViolation<T>> beanValidationFailures) {
        this.success = false;
        this.actionType = "validation";
        this.data = new HashMap();
        StringBuffer message = new StringBuffer("<p>");
        Iterator i$ = beanValidationFailures.iterator();

        while(i$.hasNext()) {
            ConstraintViolation failure = (ConstraintViolation)i$.next();
            message.append(failure.getMessage() + "<p>");
        }

        this.actionMessage = message.toString();
    }

    public boolean getSuccess() {
        return this.success;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }

    public boolean isDoRedirect() {
        return this.doRedirect;
    }

    public void setDoRedirect(boolean doRedirect) {
        this.doRedirect = doRedirect;
    }

    public String getRedirectUrl() {
        return this.redirectUrl;
    }

    public void setRedirectUrl(String redirectUrl) {
        this.redirectUrl = redirectUrl;
    }

    public String getActionMessage() {
        return this.actionMessage;
    }

    public void setActionMessage(String actionMessage) {
        this.actionMessage = actionMessage;
    }

    public Object getData() {
        return this.data;
    }

    public void setData(Object data) {
        this.data = data;
    }

    public void putData(Object data) {
        this.data = data;
    }

    public Map<String, Object> getOtherAttributes() {
        return this.otherAttributes;
    }

    public void setOtherAttributes(String key, Object otherAttribute) {
        if(this.otherAttributes == null) {
            this.otherAttributes = new HashMap();
        }

        this.otherAttributes.put(key, otherAttribute);
    }

    public void setOtherAttributes(Map<String, Object> otherAttributes) {
        this.otherAttributes = otherAttributes;
    }

    public String getActionType() {
        return this.actionType;
    }

    public void setActionType(String actionType) {
        this.actionType = actionType;
    }

    public long getTotalResult() {
        return this.totalResult;
    }

    public void setTotalResult(long totalResult) {
        this.totalResult = totalResult;
    }
}
