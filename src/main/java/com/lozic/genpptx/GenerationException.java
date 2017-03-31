package com.lozic.genpptx;

public class GenerationException extends Exception {

    private static final long serialVersionUID = -5019484515695525707L;

    public GenerationException() {
        super();
    }

    public GenerationException(String message) {
        super(message);
    }

    public GenerationException(Throwable cause) {
        super(cause);
    }

    public GenerationException(String message, Throwable cause) {
        super(message, cause);
    }

    public GenerationException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }

}
