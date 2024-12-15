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

#include "xlsxwriter.h"

using namespace std;

// globalno ustvarimo niz 'danasnji_dan'
// s pomocjo makroja __DATE__
unordered_map<string, string> mesci = {
        {"Jan", "01"}, {"Feb", "02"}, {"Mar", "03"},
        {"Apr", "04"}, {"May", "05"}, {"Jun", "06"},
        {"Jul", "07"}, {"Aug", "08"}, {"Sep", "09"},
        {"Oct", "10"}, {"Nov", "11"}, {"Dec", "12"}
    };

string danasnji_dan = __DATE__;
string dan = danasnji_dan.substr(4,2);
string mesec = mesci[danasnji_dan.substr(0,3)];
string leto = danasnji_dan.substr(7,4);

string ocisti_niz(string stavek, string vzorec, string zamenjava = " ") {
    int pos = 0;
    while ((pos = stavek.find(vzorec, pos)) != string::npos) {
        stavek.replace(pos, vzorec.length(), zamenjava);
        pos += zamenjava.length();
    }
    return stavek;
}

void FDX(const vector<vector<string>>& mat) {
    /*Funkcija nam iz parametra mat - matrike ustvari izhodno FDX datoteko*/
    string ime_dat = "VAT AND DUTIES " + dan + "." + mesec + "." + leto + " FEDEX.xlsx";
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
        for (;i < 24; ++i) {
            try {
                niz = ocisti_niz(vek[i], ".", "\0");
                st = stof(ocisti_niz(niz, ",", "."));
                // da zaokrozimo stevilo na dve decimalni mesti delimo z 100.0
                worksheet_write_number(ws, vr, i, round(st * 100.0) / 100.0, format_dec);
            } catch (invalid_argument) {
                worksheet_write_string(ws, vr, i, vek[i].c_str(), NULL);
            }
        }
        for (;i < 27; ++i) {
            worksheet_write_string(ws, vr, i, vek[i].c_str(), NULL);
        }
        ++vr;
    }

    workbook_close(wb);
}

#endif // FUNKCIJE
