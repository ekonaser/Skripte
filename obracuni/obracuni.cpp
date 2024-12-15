#include <algorithm>
#include <iostream>
#include <sstream>
#include <fstream>
#include <vector>
#include <unordered_map>
#include <set>
#include <ctime>

#include "xlsxwriter.h"

using namespace std;

//###################################FUNKCIJE###################################//
void FDX(const vector<vector<wstring>> mat) {
    /*Funkcija nam iz parametra mat - matrike ustvari izhodno FDX datoteko*/
    lxw_workbook  *wb  = workbook_new("VAT AND DUTIES FEDEX.xlsx");
    lxw_worksheet *ws = workbook_add_worksheet(wb, "FDX");
    //worksheet_write_string(worksheet, 0, 0, "Hello, Excel!", NULL);

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

    // Format za glavne celice st.
    lxw_format *format_glava_st = workbook_add_format(wb);
    format_set_align(format_glava_st, LXW_ALIGN_CENTER);
    format_set_align(format_glava_st, LXW_ALIGN_VERTICAL_CENTER);
    format_set_font_size(format_glava_st, 8);
    format_set_border(format_glava_st, LXW_BORDER_THIN);
    format_set_font_name(format_glava_st, "Calibri");
    format_set_num_format(format_glava_st, "000");

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

    worksheet_write_string(ws, 3, 0, "Con Note", format_glava);
    worksheet_write_string(ws, 3, 1, "Customer reference", format_glava);
    worksheet_write_string(ws, 3, 2, "MRN", format_glava);
    worksheet_write_string(ws, 3, 3, "Date", format_glava);
    worksheet_write_string(ws, 3, 4, "address", format_glava);
    worksheet_write_string(ws, 3, 5, "postcode", format_glava);
    worksheet_write_string(ws, 3, 6, "town", format_glava);
    worksheet_write_string(ws, 3, 7, "Account no.", format_glava);
    worksheet_write_string(ws, 3, 8, "Customer name", format_glava);
    worksheet_write_string(ws, 3, 9, "VAT ID", format_glava);
    worksheet_write_string(ws, 3, 10, "Customer name", format_glava);
    worksheet_write_string(ws, 3, 11, "59", format_glava_st);
    worksheet_write_string(ws, 3, 12, "52", format_glava_st);
    worksheet_write_string(ws, 3, 13, "425", format_glava_st);
    worksheet_write_string(ws, 3, 14, "350", format_glava_st);
    worksheet_write_string(ws, 3, 15, "422", format_glava_st);
    worksheet_write_string(ws, 3, 16, "74", format_glava_st);
    worksheet_write_string(ws, 3, 17, "429", format_glava_st);
    worksheet_write_string(ws, 3, 18, "414", format_glava_st);
    worksheet_write_string(ws, 3, 19, "415", format_glava_st);
    worksheet_write_string(ws, 3, 20, "411", format_glava_st);
    worksheet_write_string(ws, 3, 21, "424", format_glava_st);
    worksheet_write_string(ws, 3, 22, "421", format_glava_st);
    worksheet_write_string(ws, 3, 23, "404", format_glava_st);
    worksheet_write_string(ws, 3, 24, "432", format_glava_st);
    worksheet_write_string(ws, 3, 25, "940", format_glava_st);
    worksheet_write_string(ws, 3, 26, "---", format_glava);

    workbook_close(wb);
}

wstring ocisti_niz(wstring stavek, wstring vzorec, wstring zamenjava = L" ") {
    int pos = 0;
    while ((pos = stavek.find(vzorec, pos)) != wstring::npos) {
        stavek.replace(pos, vzorec.length(), zamenjava);
        pos += zamenjava.length();
    }
    return stavek;
}

vector<wstring> razdeli(wstring stavek, wchar_t del = L';') {
    /*Funkcija nam vhodni niz razdeli s pomocjo podanega locila*/
    wstringstream tok_niza(stavek);
    vector<wstring> sez_bes;
    wstring beseda;

    while (getline(tok_niza, beseda, del)) {
        sez_bes.push_back(beseda);
    }
    
    return sez_bes;
}

unordered_map<wstring, wstring> ustvariSlovar(const string datoteka, short int a = 0, short int b = 1) {
    /*Funkcija nam iz vhodne datoteke ustvari slovar*/
    unordered_map<wstring, wstring> slo;
    wifstream dat(datoteka);
    vector<wstring> vek;
    wstring vr;
    while (getline(dat, vr)) {
        vek = razdeli(vr);
        slo.insert({vek[a], vek[b]});
    }
    return slo;
}

//####################################RAZRED####################################//
class Slovar{
    public:
        Slovar(const string datoteka) {
            slo = ustvariSlovar(datoteka);
        }
    // globalne spremenljivke
    unordered_map<wstring, wstring> slo;

    // metode
        set<wstring> vrednosti() {
            /*Metoda nam vrne mnozico vrednosti*/
            set<wstring> sez_vre;
            for (const auto par: slo) {
                sez_vre.insert(par.second);
            }
            return sez_vre;
        }

        set<wstring> kljuci() {
            /*Metoda nam vrne mnozico kljucev*/
            set<wstring> sez_vre;
            for (const auto par: slo) {
                sez_vre.insert(par.first);
            }
            return sez_vre;
        }
};
//##############################################################################//

vector<vector<wstring>> izlusci(const string datoteka) {
    /*  Funkcija nam iz .xml datoteke izlusci vrstice in stolpce
        ter podatke vrne kot vektorje v vektorju - matrika. Ta
        pristop je veliko bolj 'razumljiv' od .findall() metode
        v Pythonu.
    */
    vector<vector<wstring>> podatki;
    wifstream dat(datoteka);
    wstring vr;
    while (getline(dat, vr)) {
        // poiscemo vrstico pri kateri se podatki zacnejo
        if (vr.find(L"<Row") != wstring::npos) {
            vector<wstring> vrstica;
            while (getline(dat,vr)) {
                // Dodamo samo tisto vrstico v kateri so podatki
                if (vr.find(L"<Cell ss:Index=") != wstring::npos) {
                    vrstica.push_back(L"");
                } else if (vr.find(L"<Data ss:Type=\"String\">") != wstring::npos) {
                    // -30 zaradi tega ker je 'prazna' celica dolzine 30
                    // se prej pa ocistimo niz od '\n' oz. '&#10'
                    // in od & znaka oz. &amp;
                    if (vr.find(L"&#10;") != wstring::npos) {
                        vr = ocisti_niz(vr, L"&#10;");
                    } else if (vr.find(L"&amp;") != wstring::npos) {
                        vr = ocisti_niz(vr, L"&amp;", L"&");
                    }
                    vrstica.back() = vr.substr(23,size(vr)-30);
                } else if (vr.find(L"</Row>") != wstring::npos) {
                    podatki.push_back(vrstica);
                    break;
                }
            }
        }
    }
    dat.close();
    return podatki;
}

int main() {
    wofstream dat("obracun.csv");
    // slovar parov stevilk AWB:MRN
    Slovar slovar_parov("carm_scan.csv");
    // slovar acc numbers
    unordered_map<wstring, wstring> slo_acc_num = ustvariSlovar("navadna_acc.csv", 1, 0);
    // matrika podatkov
    vector<vector<wstring>> mat_pod;
    vector<vector<wstring>> mat_fdx;
    vector<vector<wstring>> mat_tnt;
    mat_pod = izlusci("obracun_tlm.xls");
    
    for (vector vek : mat_pod) {
        if (size(vek[0]) > 12) {
            vek[0].erase(remove(vek[0].begin(), vek[0].end(), ' '), vek[0].end());
            vek[0] = razdeli(vek[0],'-')[1];
        }
        if (slovar_parov.kljuci().count(vek[0]) > 0 || 
            slovar_parov.vrednosti().count(vek[2]) > 0 ||
            vek[0] == L"FEDEX" || vek[0] == L"TNT") {
                if (slo_acc_num.count(vek[9]) > 0 && size(vek[0]) == 11) {
                    vek[8] = slo_acc_num[vek[9]];
                }
                if (vek[7] == L"") {
                    vek[7] = L"CPT";
                }
                if (size(vek[0]) == 12) {
                    mat_fdx.push_back(vek);
                }
                if (size(vek[0]) == 11) {
                    mat_tnt.push_back(vek);
                }
                for (wstring celica : vek) {
                    dat << celica << ";";
                }
                dat << "\n";
        }
    }
    dat.close();
    FDX(mat_fdx);
    return 0;
}
