package com.apex.reconciliation_app.exception;

public class BucketOverflowException extends RuntimeException{
    public BucketOverflowException(String message) {
        super(message);
    }
}
