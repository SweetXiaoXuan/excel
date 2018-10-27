package com.excel.easyexcel;

public class BookModel {
    private String bookName;
    /** 作者信息 */
    private AuthorModel author;

    public String getBookName() {
        return bookName;
    }

    public void setBookName(String bookName) {
        this.bookName = bookName;
    }

    public AuthorModel getAuthor() {
        return author;
    }

    public void setAuthor(AuthorModel author) {
        this.author = author;
    }
}
