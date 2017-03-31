package com.lozic.genpptx.css;

/**
 * Simple aggregator for css style.
 *
 */
public class Style {

    private boolean bold;
    private boolean italic;
    private boolean underline;
    private String color;
    private int fontSize;
    private String liColor;
    private String liChar;

    public boolean isBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public boolean isItalic() {
        return italic;
    }

    public void setItalic(boolean italic) {
        this.italic = italic;
    }

    public boolean isUnderline() {
        return underline;
    }

    public void setUnderline(boolean underline) {
        this.underline = underline;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public int getFontSize() {
        return fontSize;
    }

    public void setFontSize(int fontSize) {
        this.fontSize = fontSize;
    }

    public String getLiColor() {
        return liColor;
    }

    public void setLiColor(String liColor) {
        this.liColor = liColor;
    }

    public String getLiChar() {
        return liChar;
    }

    public void setLiChar(String liChar) {
        this.liChar = liChar;
    }

}