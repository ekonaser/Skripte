#ifndef FUNCTIONS_H
#define FUNCTIONS_H

#include <vector>
#include <string>
#include <algorithm>
#include <iostream>
#include <sstream>
#include <fstream>
#include <unordered_map>
#include <set>
#include <cmath>
#include <cstdio>

#include "xlsxwriter.h"

using namespace std;

string ocisti_niz(string stavek, string vzorec, string zamenjava = " ") {
    int pos = 0;
    while ((pos = stavek.find(vzorec, pos)) != string::npos) {
        stavek.replace(pos, vzorec.length(), zamenjava);
        pos += zamenjava.length();
    }
    return stavek;
}

int FDX(const vector<vector<string>>& mat) {
    /*Funkcija nam iz parametra mat - matrike ustvari izhodno FDX datoteko
        Funkcija vrne vrednost 0 v vsakem primeru.
    */
    if (mat.empty()) {
        return 0;
    }
    string ime_dat = "VAT AND DUTIES DD.MM.LLLL FEDEX.xlsx";
    lxw_workbook  *wb  = workbook_new(ime_dat.c_str());
    lxw_worksheet *ws = workbook_add_worksheet(wb, "FDX");

    // Format za FedEx naslov
    lxw_format *format_fdx = workbook_add_format(wb);
    format_set_bg_color(format_fdx, 0x4D148C);
    format_set_bold(format_fdx);
    format_set_align(format_fdx, LXW_ALIGN_CENTER);
    format_set_align(format_fdx, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_name(format_fdx, "Calibri");
    format_set_font_size(format_fdx, 36);
    format_set_border(format_fdx, LXW_BORDER_THIN);
    format_set_font_color(format_fdx, 0xFFFFFF);

    // Format za glavne celice
    lxw_format *format_glava = workbook_add_format(wb);
    format_set_align(format_glava, LXW_ALIGN_CENTER);
    format_set_align(format_glava, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_size(format_glava, 8);
    format_set_border(format_glava, LXW_BORDER_THIN);
    format_set_font_name(format_glava, "Calibri");
    format_set_text_wrap(format_glava);

    // Format za glavne celice 2
    lxw_format *format_glava2 = workbook_add_format(wb);
    format_set_align(format_glava2, LXW_ALIGN_CENTER);
    format_set_align(format_glava2, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_size(format_glava2, 8);
    format_set_border(format_glava2, LXW_BORDER_THIN);
    format_set_font_name(format_glava2, "Calibri");

    // Navaden format
    lxw_format *format = workbook_add_format(wb);
    format_set_font_size(format, 11);
    format_set_font_name(format, "Calibri");

    // Format za glavne celice st.
    lxw_format *format_glava_st = workbook_add_format(wb);
    format_set_align(format_glava_st, LXW_ALIGN_CENTER);
    format_set_align(format_glava_st, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_size(format_glava_st, 8);
    format_set_border(format_glava_st, LXW_BORDER_THIN);
    format_set_font_name(format_glava_st, "Calibri");
    format_set_num_format(format_glava_st, "000");

    // Format za cela stevila
    lxw_format *format_int = workbook_add_format(wb);
    format_set_font_name(format_int, "Calibri");
    format_set_num_format(format_int, "0");

    // Format za decimalna stevila
    lxw_format *format_dec = workbook_add_format(wb);
    format_set_font_name(format_dec, "Calibri");
    format_set_num_format(format_dec, "0.00");

    // zdruzimo celice
    worksheet_merge_range(ws, 2, 0, 2, 9, "FedEx", format_fdx);
    worksheet_set_row(ws, 2, 100, NULL);
    
    worksheet_write_string(ws, 2, 10, "Description", format_glava);
    worksheet_write_string(ws, 2, 11, "VAT", format_glava);
    worksheet_write_string(ws, 2, 12, "Duties", format_glava);
    worksheet_write_string(ws, 2, 13, "Other\nGovernment\nAgency", format_glava);
    worksheet_write_string(ws, 2, 14, "Additional\nLine\nItems", format_glava);
    worksheet_write_string(ws, 2, 15, "Clearence\ntransfer", format_glava);
    worksheet_write_string(ws, 2, 16, "Disbursement\nFee", format_glava);
    worksheet_write_string(ws, 2, 17, "Pre-Payment\n(Direct\nPayment\nProcessing)", format_glava);
    worksheet_write_string(ws, 2, 18, "Returned\nGoods", format_glava);
    worksheet_write_string(ws, 2, 19, "Temporary\nImport", format_glava);
    worksheet_write_string(ws, 2, 20, "Post Entry\nAdjustment", format_glava);
    worksheet_write_string(ws, 2, 21, "In-Bond\nTransit", format_glava);
    worksheet_write_string(ws, 2, 22, "Storage", format_glava);
    worksheet_write_string(ws, 2, 23, "Customized\nService", format_glava);
    worksheet_write_string(ws, 2, 24, "Brokerage\nFee", format_glava);
    worksheet_write_string(ws, 2, 25, "VAT on\nAncillary\nFees (@\n22%)", format_glava);
    worksheet_write_string(ws, 2, 26, "Total\nCollected\nfrom\nCustomer", format_glava);

    worksheet_write_string(ws, 3, 0, "Con Note", format_glava2);
    worksheet_write_string(ws, 3, 1, "Customer reference", format_glava2);
    worksheet_write_string(ws, 3, 2, "MRN", format_glava2);
    worksheet_write_string(ws, 3, 3, "Date", format_glava2);
    worksheet_write_string(ws, 3, 4, "address", format_glava2);
    worksheet_write_string(ws, 3, 5, "postcode", format_glava2);
    worksheet_write_string(ws, 3, 6, "town", format_glava2);
    worksheet_write_string(ws, 3, 7, "Account no.", format_glava2);
    worksheet_write_string(ws, 3, 8, "Customer name", format_glava2);
    worksheet_write_string(ws, 3, 9, "VAT ID", format_glava2);
    worksheet_write_string(ws, 3, 10, "Customer name", format_glava2);
    worksheet_write_number(ws, 3, 11, 59, format_glava_st);
    worksheet_write_number(ws, 3, 12, 52, format_glava_st);
    worksheet_write_number(ws, 3, 13, 425, format_glava_st);
    worksheet_write_number(ws, 3, 14, 350, format_glava_st);
    worksheet_write_number(ws, 3, 15, 422, format_glava_st);
    worksheet_write_number(ws, 3, 16, 74, format_glava_st);
    worksheet_write_number(ws, 3, 17, 429, format_glava_st);
    worksheet_write_number(ws, 3, 18, 414, format_glava_st);
    worksheet_write_number(ws, 3, 19, 415, format_glava_st);
    worksheet_write_number(ws, 3, 20, 411, format_glava_st);
    worksheet_write_number(ws, 3, 21, 424, format_glava_st);
    worksheet_write_number(ws, 3, 22, 421, format_glava_st);
    worksheet_write_number(ws, 3, 23, 404, format_glava_st);
    worksheet_write_number(ws, 3, 24, 432, format_glava_st);
    worksheet_write_number(ws, 3, 25, 940, format_glava_st);
    worksheet_write_string(ws, 3, 26, "---", format_glava2);

    worksheet_autofilter(ws, 3, 0, 3, 26);

    int vr = 4;
    string niz;
    float st;
    for(vector<string> vek : mat) {
        int i = 1;
        if (vek[7] == "DDP") {
            vek[8] = "SHIPPER";
        } else {
            vek[8] = "RECEIVER";
        }
        try {
            worksheet_write_number(ws, vr, 0, stoll(vek[0]), format_int);
        } catch (invalid_argument) {
            worksheet_write_string(ws, vr, 0, vek[0].c_str(), NULL);
        }
        for (;i < 11; ++i) {
            worksheet_write_string(ws, vr, i, vek[i].c_str(), NULL);
        }
        
        float skupaj = 0;
        for (;i < 24; ++i) {
            try {
                niz = ocisti_niz(vek[i], ".", "\0");
                st = stof(ocisti_niz(niz, ",", "."));
                // da zaokrozimo stevilo na dve decimalni mesti delimo z 100.0
                skupaj += st;
                worksheet_write_number(ws, vr, i, round(st * 100.0) / 100.0, format_dec);
            } catch (invalid_argument) {
                worksheet_write_string(ws, vr, i, vek[i].c_str(), NULL);
            }
        }
        for (;i < 27; ++i) {
            worksheet_write_string(ws, vr, i, vek[i].c_str(), NULL);
        }
        if (vek[24] == "ROD") {
            niz = ocisti_niz(vek[16], ".", "\0");
            st = stof(ocisti_niz(niz, ",", ".")) * 0.22;
            worksheet_write_number(ws, vr, 25, round(st * 100.0) / 100.0, format_dec);
            worksheet_write_number(ws, vr, 26, round((skupaj + st) * 100.0) / 100.0, format_dec);
        }
        ++vr;
    }

    workbook_close(wb);
    return 0;
}

int TNT_NAV(const vector<vector<string>>& mat) {
    /*Funkcija nam iz parametra mat - matrike ustvari izhodno TNT NAV datoteko
        Funkcija vrne vrednost 0 v vsakem primeru.
    */
    if (mat.empty()) {
        return 0;
    }
    string ime_dat = "VAT AND DUTIES DD.MM.LLLL NAVADNA.xlsx";
    lxw_workbook  *wb  = workbook_new(ime_dat.c_str());
    lxw_worksheet *ws = workbook_add_worksheet(wb, "TNT");

    // Format za TNT naslov
    lxw_format *format_tnt = workbook_add_format(wb);
    format_set_bg_color(format_tnt, 0xFF6200);
    format_set_bold(format_tnt);
    format_set_align(format_tnt, LXW_ALIGN_CENTER);
    format_set_align(format_tnt, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_name(format_tnt, "Calibri");
    format_set_font_size(format_tnt, 36);
    format_set_border(format_tnt, LXW_BORDER_THIN);
    format_set_font_color(format_tnt, 0xFFFFFF);

    // Format za glavne celice
    lxw_format *format_glava = workbook_add_format(wb);
    format_set_align(format_glava, LXW_ALIGN_CENTER);
    format_set_align(format_glava, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_size(format_glava, 8);
    format_set_border(format_glava, LXW_BORDER_THIN);
    format_set_font_name(format_glava, "Calibri");
    format_set_text_wrap(format_glava);

    // Format za glavne celice 2
    lxw_format *format_glava2 = workbook_add_format(wb);
    format_set_align(format_glava2, LXW_ALIGN_CENTER);
    format_set_align(format_glava2, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_size(format_glava2, 8);
    format_set_border(format_glava2, LXW_BORDER_THIN);
    format_set_font_name(format_glava2, "Calibri");

    // Format za cela stevila
    lxw_format *format_int = workbook_add_format(wb);
    format_set_font_name(format_int, "Calibri");
    format_set_num_format(format_int, "0");

    // Format za decimalna stevila
    lxw_format *format_dec = workbook_add_format(wb);
    format_set_font_name(format_dec, "Calibri");
    format_set_num_format(format_dec, "0.00");

    worksheet_merge_range(ws, 2, 0, 2, 6, "TNT", format_tnt);
    worksheet_set_row(ws, 2, 100, NULL);

    worksheet_write_string(ws, 2, 7, "Description", format_glava);
    worksheet_write_string(ws, 2, 8, "VAT", format_glava);
    worksheet_write_string(ws, 2, 9, "Duties", format_glava);
    worksheet_write_string(ws, 2, 10, "Other\nGovernment\nAgency", format_glava);
    worksheet_write_string(ws, 2, 11, "Additional\nLine Items", format_glava);
    worksheet_write_string(ws, 2, 12, "Clearence transfer", format_glava);
    worksheet_write_string(ws, 2, 13, "Disbursement Fee\nReturned Goods\nPre-Payment\n(Direct Payment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment", format_glava);
    worksheet_write_string(ws, 2, 14, "In-Bond Transit", format_glava);
    worksheet_write_string(ws, 2, 15, "Storage\nCustomized\nService", format_glava);
    worksheet_write_string(ws, 2, 16, "Disbursement Fee\nReturned Goods\nPre-Payment\n(Direct Payment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment", format_glava);
    worksheet_write_string(ws, 2, 17, "Clearence transfer", format_glava);
    worksheet_write_string(ws, 2, 18, "Additional\nLine\nItems", format_glava);
    worksheet_write_string(ws, 2, 19, "Storage\nCustomized\nService", format_glava);
    worksheet_write_string(ws, 2, 20, "In-Bond Transit", format_glava);
    worksheet_write_string(ws, 2, 21, "In-Bond Transit", format_glava);
    worksheet_write_string(ws, 2, 22, "Storage\nCustomized\nService", format_glava);
    worksheet_write_string(ws, 2, 23, "Disbursement Fee\nReturned Goods\nPre-Payment\n(Direct Payment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment", format_glava);
    worksheet_write_string(ws, 2, 24, "Clearence transfer", format_glava);

    worksheet_write_string(ws, 3, 0, "Con Note", format_glava2);
    worksheet_write_string(ws, 3, 1, "Customer reference", format_glava2);
    worksheet_write_string(ws, 3, 2, "MRN", format_glava2);
    worksheet_write_string(ws, 3, 3, "Date", format_glava2);
    worksheet_write_string(ws, 3, 4, "Account no.", format_glava2);
    worksheet_write_string(ws, 3, 5, "Customer name", format_glava2);
    worksheet_write_string(ws, 3, 6, "VAT ID", format_glava2);
    worksheet_write_string(ws, 3, 7, "Customer name", format_glava2);
    worksheet_write_string(ws, 3, 8, "VT", format_glava2);
    worksheet_write_string(ws, 3, 9, "DT", format_glava2);
    worksheet_write_string(ws, 3, 10, "CC1", format_glava2);
    worksheet_write_string(ws, 3, 11, "CL0", format_glava2);
    worksheet_write_string(ws, 3, 12, "CL3", format_glava2);
    worksheet_write_string(ws, 3, 13, "CL5", format_glava2);
    worksheet_write_string(ws, 3, 14, "CL6", format_glava2);
    worksheet_write_string(ws, 3, 15, "HF", format_glava2);
    worksheet_write_string(ws, 3, 16, "CG1", format_glava2);
    worksheet_write_string(ws, 3, 17, "CL4", format_glava2);
    worksheet_write_string(ws, 3, 18, "OCH", format_glava2);
    worksheet_write_string(ws, 3, 19, "ST1", format_glava2);
    worksheet_write_string(ws, 3, 20, "TR1", format_glava2);
    worksheet_write_string(ws, 3, 21, "CL9", format_glava2);
    worksheet_write_string(ws, 3, 22, "CLA", format_glava2);
    worksheet_write_string(ws, 3, 23, "CLB", format_glava2);
    worksheet_write_string(ws, 3, 24, "CLC", format_glava2);

    worksheet_autofilter(ws, 3, 0, 3, 24);

    int vr = 4;
    string niz;
    float st;
    for(vector<string> vek : mat) {
        int i = 1;
        int x = 8;
        try {
            worksheet_write_number(ws, vr, 0, stoll(vek[0]), format_int);
        } catch (invalid_argument) {
            worksheet_write_string(ws, vr, 0, vek[0].c_str(), NULL);
        }
        for (;i < 5; ++i) {
            worksheet_write_string(ws, vr, i, vek[i].c_str(), NULL);
        }
        worksheet_write_string(ws, vr, 4, vek[8].c_str(), NULL);
        worksheet_write_string(ws, vr, 5, vek[7].c_str(), NULL);
        worksheet_write_string(ws, vr, 6, vek[9].c_str(), NULL);
        worksheet_write_string(ws, vr, 7, vek[10].c_str(), NULL);
        i = 11;
        for (;x < 25; ++x) {
            try {
                niz = ocisti_niz(vek[i], ".", "\0");
                st = stof(ocisti_niz(niz, ",", "."));
                // da zaokrozimo stevilo na dve decimalni mesti delimo z 100.0
                worksheet_write_number(ws, vr, x, round(st * 100.0) / 100.0, format_dec);
            } catch (invalid_argument) {
                worksheet_write_string(ws, vr, x, vek[i].c_str(), NULL);
            }
            ++i;
        }
        ++vr;
    }

    workbook_close(wb);
    return 0;
}

int TNT_ISB(const vector<vector<string>>& mat) {
    /*Funkcija nam iz parametra mat - matrike ustvari izhodno TNT ISB datoteko
        Funkcija vrne vrednost 0 v vsakem primeru.
    */
    if (mat.empty()) {
        return 0;
    }
    string ime_dat = "VAT AND DUTIES DD.MM.LLLL ISB OUT.xlsx";
    lxw_workbook  *wb  = workbook_new(ime_dat.c_str());
    lxw_worksheet *ws = workbook_add_worksheet(wb, "TNT");

    // Format za TNT naslov
    lxw_format *format_tnt = workbook_add_format(wb);
    format_set_bg_color(format_tnt, 0xFF6200);
    format_set_bold(format_tnt);
    format_set_align(format_tnt, LXW_ALIGN_CENTER);
    format_set_align(format_tnt, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_name(format_tnt, "Calibri");
    format_set_font_size(format_tnt, 36);
    format_set_border(format_tnt, LXW_BORDER_THIN);
    format_set_font_color(format_tnt, 0xFFFFFF);

    // Format za glavne celice
    lxw_format *format_glava = workbook_add_format(wb);
    format_set_align(format_glava, LXW_ALIGN_CENTER);
    format_set_align(format_glava, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_size(format_glava, 8);
    format_set_border(format_glava, LXW_BORDER_THIN);
    format_set_font_name(format_glava, "Calibri");
    format_set_text_wrap(format_glava);

    // Format za glavne celice 2
    lxw_format *format_glava2 = workbook_add_format(wb);
    format_set_align(format_glava2, LXW_ALIGN_CENTER);
    format_set_align(format_glava2, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_size(format_glava2, 8);
    format_set_border(format_glava2, LXW_BORDER_THIN);
    format_set_font_name(format_glava2, "Calibri");

    // Format za cela stevila
    lxw_format *format_int = workbook_add_format(wb);
    format_set_font_name(format_int, "Calibri");
    format_set_num_format(format_int, "0");

    // Format za decimalna stevila
    lxw_format *format_dec = workbook_add_format(wb);
    format_set_font_name(format_dec, "Calibri");
    format_set_num_format(format_dec, "0.00");

    worksheet_merge_range(ws, 2, 0, 2, 7, "TNT", format_tnt);
    worksheet_set_row(ws, 2, 100, NULL);

    worksheet_write_string(ws, 2, 8, "Description", format_glava);
    worksheet_write_string(ws, 2, 9, "VAT", format_glava);
    worksheet_write_string(ws, 2, 10, "Duties", format_glava);
    worksheet_write_string(ws, 2, 11, "Other\nGovernment\nAgency", format_glava);
    worksheet_write_string(ws, 2, 12, "Additional\nLine Items", format_glava);
    worksheet_write_string(ws, 2, 13, "Clearence transfer", format_glava);
    worksheet_write_string(ws, 2, 14, "Disbursement Fee\nReturned Goods\nPre-Payment\n(Direct Payment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment", format_glava);
    worksheet_write_string(ws, 2, 15, "In-Bond Transit", format_glava);
    worksheet_write_string(ws, 2, 16, "Storage\nCustomized\nService", format_glava);
    worksheet_write_string(ws, 2, 17, "Disbursement Fee\nReturned Goods\nPre-Payment\n(Direct Payment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment", format_glava);
    worksheet_write_string(ws, 2, 18, "Clearence transfer", format_glava);
    worksheet_write_string(ws, 2, 19, "Additional\nLine\nItems", format_glava);
    worksheet_write_string(ws, 2, 20, "Storage\nCustomized\nService", format_glava);
    worksheet_write_string(ws, 2, 21, "In-Bond Transit", format_glava);
    worksheet_write_string(ws, 2, 22, "In-Bond Transit", format_glava);
    worksheet_write_string(ws, 2, 23, "Storage\nCustomized\nService", format_glava);
    worksheet_write_string(ws, 2, 24, "Disbursement Fee\nReturned Goods\nPre-Payment\n(Direct Payment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment", format_glava);
    worksheet_write_string(ws, 2, 25, "Clearence transfer", format_glava);

    worksheet_write_string(ws, 3, 0, "Con Note", format_glava2);
    worksheet_write_string(ws, 3, 1, "Customer reference", format_glava2);
    worksheet_write_string(ws, 3, 2, "MRN", format_glava2);
    worksheet_write_string(ws, 3, 3, "Date", format_glava2);
    worksheet_write_string(ws, 3, 4, "Account no.", format_glava2);
    worksheet_write_string(ws, 3, 5, "Customer name", format_glava2);
    worksheet_write_string(ws, 3, 6, "product", format_glava2);
    worksheet_write_string(ws, 3, 7, "division", format_glava2);
    worksheet_write_string(ws, 3, 8, "Customer name", format_glava2);
    worksheet_write_string(ws, 3, 9, "VT", format_glava2);
    worksheet_write_string(ws, 3, 10, "DT", format_glava2);
    worksheet_write_string(ws, 3, 11, "CC1", format_glava2);
    worksheet_write_string(ws, 3, 12, "CL0", format_glava2);
    worksheet_write_string(ws, 3, 13, "CL3", format_glava2);
    worksheet_write_string(ws, 3, 14, "CL5", format_glava2);
    worksheet_write_string(ws, 3, 15, "CL6", format_glava2);
    worksheet_write_string(ws, 3, 16, "HF", format_glava2);
    worksheet_write_string(ws, 3, 17, "CG1", format_glava2);
    worksheet_write_string(ws, 3, 18, "CL4", format_glava2);
    worksheet_write_string(ws, 3, 19, "OCH", format_glava2);
    worksheet_write_string(ws, 3, 20, "ST1", format_glava2);
    worksheet_write_string(ws, 3, 21, "TR1", format_glava2);
    worksheet_write_string(ws, 3, 22, "CL9", format_glava2);
    worksheet_write_string(ws, 3, 23, "CLA", format_glava2);
    worksheet_write_string(ws, 3, 24, "CLB", format_glava2);
    worksheet_write_string(ws, 3, 25, "CLC", format_glava2);

    worksheet_autofilter(ws, 3, 0, 3, 25);

    int vr = 4;
    string niz;
    float st;
    for (vector<string> vek : mat) {
        int i = 1;
        int x = 9;
        try {
            worksheet_write_number(ws, vr, 0, stoll(vek[0]), format_int);
        } catch (invalid_argument) {
            worksheet_write_string(ws, vr, 0, vek[0].c_str(), NULL);
        }
        for (;i < 4; ++i) {
            worksheet_write_string(ws, vr, i, vek[i].c_str(), NULL);
        }
        worksheet_write_string(ws, vr, 4, "DDP", NULL);
        worksheet_write_string(ws, vr, 6, "Z", NULL);
        worksheet_write_number(ws, vr, 7, 105, format_int);
        i = 11;
        for (;x < 25; ++x) {
            try {
                niz = ocisti_niz(vek[i], ".", "\0");
                st = stof(ocisti_niz(niz, ",", "."));
                // da zaokrozimo stevilo na dve decimalni mesti delimo z 100.0
                worksheet_write_number(ws, vr, x, round(st * 100.0) / 100.0, format_dec);
            } catch (invalid_argument) {
                worksheet_write_string(ws, vr, x, vek[i].c_str(), NULL);
            }
            ++i;
        }
        ++vr;
    }

    workbook_close(wb);
    return 0;
}

int TNT_PRIVATE(const vector<vector<string>>& mat) {
    /*Funkcija nam iz parametra mat - matrike ustvari izhodno TNT ISB datoteko
        Funkcija vrne vrednost 0 v vsakem primeru.
    */
    if (mat.empty()) {
        return 0;
    }
    string ime_dat = "VAT AND DUTIES DD.MM.LLLL PRIVATE INDIVIDUALS.xlsx";
    lxw_workbook  *wb  = workbook_new(ime_dat.c_str());
    lxw_worksheet *ws = workbook_add_worksheet(wb, "TNT");

    // Format za TNT naslov
    lxw_format *format_tnt = workbook_add_format(wb);
    format_set_bg_color(format_tnt, 0xFF6200);
    format_set_bold(format_tnt);
    format_set_align(format_tnt, LXW_ALIGN_CENTER);
    format_set_align(format_tnt, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_name(format_tnt, "Calibri");
    format_set_font_size(format_tnt, 36);
    format_set_border(format_tnt, LXW_BORDER_THIN);
    format_set_font_color(format_tnt, 0xFFFFFF);

    // Format za glavne celice
    lxw_format *format_glava = workbook_add_format(wb);
    format_set_align(format_glava, LXW_ALIGN_CENTER);
    format_set_align(format_glava, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_size(format_glava, 8);
    format_set_border(format_glava, LXW_BORDER_THIN);
    format_set_font_name(format_glava, "Calibri");
    format_set_text_wrap(format_glava);

    // Format za glavne celice 2
    lxw_format *format_glava2 = workbook_add_format(wb);
    format_set_align(format_glava2, LXW_ALIGN_CENTER);
    format_set_align(format_glava2, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_size(format_glava2, 8);
    format_set_border(format_glava2, LXW_BORDER_THIN);
    format_set_font_name(format_glava2, "Calibri");

    // Format za cela stevila
    lxw_format *format_int = workbook_add_format(wb);
    format_set_font_name(format_int, "Calibri");
    format_set_num_format(format_int, "0");

    // Format za decimalna stevila
    lxw_format *format_dec = workbook_add_format(wb);
    format_set_font_name(format_dec, "Calibri");
    format_set_num_format(format_dec, "0.00");

    worksheet_merge_range(ws, 2, 0, 2, 9, "TNT", format_tnt);
    worksheet_set_row(ws, 2, 100, NULL);

    worksheet_write_string(ws, 2, 10, "Description", format_glava);
    worksheet_write_string(ws, 2, 11, "VAT", format_glava);
    worksheet_write_string(ws, 2, 12, "Duties", format_glava);
    worksheet_write_string(ws, 2, 13, "Other\nGovernment\nAgency", format_glava);
    worksheet_write_string(ws, 2, 14, "Additional\nLine Items", format_glava);
    worksheet_write_string(ws, 2, 15, "Clearence transfer", format_glava);
    worksheet_write_string(ws, 2, 16, "Disbursement Fee\nReturned Goods\nPre-Payment\n(Direct Payment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment", format_glava);
    worksheet_write_string(ws, 2, 17, "In-Bond Transit", format_glava);
    worksheet_write_string(ws, 2, 18, "Storage\nCustomized\nService", format_glava);
    worksheet_write_string(ws, 2, 19, "Disbursement Fee\nReturned Goods\nPre-Payment\n(Direct Payment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment", format_glava);
    worksheet_write_string(ws, 2, 20, "Clearence transfer", format_glava);
    worksheet_write_string(ws, 2, 21, "Additional\nLine\nItems", format_glava);
    worksheet_write_string(ws, 2, 22, "Storage\nCustomized\nService", format_glava);
    worksheet_write_string(ws, 2, 23, "In-Bond Transit", format_glava);
    worksheet_write_string(ws, 2, 24, "In-Bond Transit", format_glava);
    worksheet_write_string(ws, 2, 25, "Storage\nCustomized\nService", format_glava);
    worksheet_write_string(ws, 2, 26, "Disbursement Fee\nReturned Goods\nPre-Payment\n(Direct Payment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment", format_glava);
    worksheet_write_string(ws, 2, 27, "Clearence transfer", format_glava);

    worksheet_write_string(ws, 3, 0, "Con Note", format_glava2);
    worksheet_write_string(ws, 3, 1, "Customer reference", format_glava2);
    worksheet_write_string(ws, 3, 2, "MRN", format_glava2);
    worksheet_write_string(ws, 3, 3, "Date", format_glava2);
    worksheet_write_string(ws, 3, 4, "address", format_glava2);
    worksheet_write_string(ws, 3, 5, "postcode", format_glava2);
    worksheet_write_string(ws, 3, 6, "town", format_glava2);
    worksheet_write_string(ws, 3, 7, "Account no.", format_glava2);
    worksheet_write_string(ws, 3, 8, "Customer name", format_glava2);
    worksheet_write_string(ws, 3, 9, "VAT ID", format_glava2);
    worksheet_write_string(ws, 3, 10, "Customer name", format_glava2);
    worksheet_write_string(ws, 3, 11, "VT", format_glava2);
    worksheet_write_string(ws, 3, 12, "DT", format_glava2);
    worksheet_write_string(ws, 3, 13, "CC1", format_glava2);
    worksheet_write_string(ws, 3, 14, "CL0", format_glava2);
    worksheet_write_string(ws, 3, 15, "CL3", format_glava2);
    worksheet_write_string(ws, 3, 16, "CL5", format_glava2);
    worksheet_write_string(ws, 3, 17, "CL6", format_glava2);
    worksheet_write_string(ws, 3, 18, "HF", format_glava2);
    worksheet_write_string(ws, 3, 19, "CG1", format_glava2);
    worksheet_write_string(ws, 3, 20, "CL4", format_glava2);
    worksheet_write_string(ws, 3, 21, "OCH", format_glava2);
    worksheet_write_string(ws, 3, 22, "ST1", format_glava2);
    worksheet_write_string(ws, 3, 23, "TR1", format_glava2);
    worksheet_write_string(ws, 3, 24, "CL9", format_glava2);
    worksheet_write_string(ws, 3, 25, "CLA", format_glava2);
    worksheet_write_string(ws, 3, 26, "CLB", format_glava2);
    worksheet_write_string(ws, 3, 27, "CLC", format_glava2);

    worksheet_autofilter(ws, 3, 0, 3, 27);

    int vr = 4;
    string niz;
    float st;
    for (vector<string> vek : mat) {
        int i = 1;
        int x = 11;
        try {
            worksheet_write_number(ws, vr, 0, stoll(vek[0]), format_int);
        } catch (invalid_argument) {
            worksheet_write_string(ws, vr, 0, vek[0].c_str(), NULL);
        }
        for (;i < 7; ++i) {
            worksheet_write_string(ws, vr, i, vek[i].c_str(), NULL);
        }
        
        worksheet_write_number(ws, vr, 7, 24914, format_int);
        worksheet_write_string(ws, vr, 8, vek[7].c_str(), NULL);
        worksheet_write_string(ws, vr, 9, vek[9].c_str(), NULL);
        worksheet_write_string(ws, vr, 10, vek[10].c_str(), NULL);

        i = 11;
        for (;x < 25; ++x) {
            try {
                niz = ocisti_niz(vek[i], ".", "\0");
                st = stof(ocisti_niz(niz, ",", "."));
                // da zaokrozimo stevilo na dve decimalni mesti delimo z 100.0
                worksheet_write_number(ws, vr, x, round(st * 100.0) / 100.0, format_dec);
            } catch (invalid_argument) {
                worksheet_write_string(ws, vr, x, vek[i].c_str(), NULL);
            }
            ++i;
        }
        ++vr;
    }

    workbook_close(wb);
    return 0;
}
#endif // FUNKCIJE
