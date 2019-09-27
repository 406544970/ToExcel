package com.example.demo.entity;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

public class ExcelTitle {
    private String name;//标题名称
    private HorizontalAlignment alignment;//对齐

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public HorizontalAlignment getAlignment() {
        return alignment;
    }

    public void setAlignment(HorizontalAlignment alignment) {
        this.alignment = alignment;
    }

    public ExcelTitle(String name, HorizontalAlignment alignment) {
        this.name = name;
        this.alignment = alignment;
    }
}
