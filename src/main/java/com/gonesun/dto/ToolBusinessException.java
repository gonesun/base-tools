package com.gonesun.dto;

public class ToolBusinessException extends RuntimeException {
    private static final long serialVersionUID = 1156978725554982429L;

    private final String code;
    private Exception innerException = null;
    private Object data;
    private Type type;

    public enum Type {
        //对应于前端的红叉提示
        error,
        //对应于前端的黄叹号提示
        warning
    }

    public ToolBusinessException(String code, String message) {
        super(message);
        this.code = code;
        this.data = null;
    }

    public ToolBusinessException(String code, String message, Type type) {
        super(message);
        this.code = code;
        this.data = null;
        this.type = type;
    }

    public ToolBusinessException(String code, String message, Object data, Exception ex) {

        super(message);
        this.code = code;
        this.data = data;
        this.innerException = ex;
    }

    public Exception getInnerException() {
        return innerException;
    }

    public Object getData() {
        return data;
    }

    public ToolBusinessException setData(Object data) {
        this.data = data;
        return this;
    }

    public ToolBusinessException setInnerException(Exception ex) {
        this.innerException = ex;
        return this;
    }

    public String getCode() {
        return code;
    }

    public Type getType() {
        return type;
    }

    public void setType(Type type) {
        this.type = type;
    }
}
