#include "stdafx.h"
#include <tchar.h>
#include <Windows.h>
#include <string.h>
#include <iostream>
#include <regex>
#include "EASendMailObj.tlh"
#include <list>
#include <sys/types.h>
#include <sys/stat.h>

#define _tprint wprint

using namespace EASendMailObjLib;

struct stat info;

const int ConnectNormal = 0;
const int ConnectSSLAuto = 1;
const int ConnectSTARTTLS = 2;
const int ConnectDirectSSL = 3;
const int ConnectTryTLS = 4;

bool Email_check(string email)
{
    const regex pattern("(\\w+)(\\.|_)?(\\w*)@(\\w+)(\\.(\\w+))+");
    if (regex_match(email, pattern))
        cout << "\nYour Email-Id is valid, you can continue\n";
    else
        cout << "\nYour Email-Id is invalid, please correct it\n";
    return regex_match(email, pattern);
}

bool Dir_check(string dir) {
    if (stat(dir.c_str(), &info) != 0)
        printf("cannot access %s\n", dir);
    else if (info.st_mode & S_IFDIR) {
        printf("%s is a directory\n", dir);
        return true;
    }
    else
        printf("%s is no directory\n", dir);
    return false;
}

void Send_mail(string mail, string msg, string to, string title, string pwd, list <string> file) {
    ::CoInitialize(NULL);
    IMailPtr oSmtp = NULL;
    oSmtp.CreateInstance(__uuidof(EASendMailObjLib::Mail));
    oSmtp->LicenseCode = _T("TryIt");
    oSmtp->FromAddr = mail.c_str();
    // Add recipient email address
    oSmtp->AddRecipientEx(to.c_str(), 0);

    // Set email subject
    oSmtp->Subject = title.c_str();
    // Set email body
    oSmtp->BodyText = msg.c_str();


    list <string> ::iterator it;
    for (it = file.begin(); it != file.end(); ++it) {
        string attachment = *it;
        Dir_check(attachment);
        if (oSmtp->AddAttachment(attachment.c_str()))
        {
            _tprintf(_T("Failed to add attachment with error: %s\r\n"), (const TCHAR*)oSmtp->GetLastErrDescription());
        }
    }
    // Your SMTP server address
    oSmtp->ServerAddr = _T("smtp.gmail.com");

    // User and password for ESMTP authentication, if your server doesn't
    // require User authentication, please remove the following codes.
    oSmtp->UserName = mail.c_str();
    oSmtp->Password = pwd.c_str();
    // Most mordern SMTP servers require SSL/TLS connection now.
    // ConnectTryTLS means if server supports SSL/TLS, SSL/TLS will be used automatically.
    oSmtp->ConnectType = ConnectTryTLS;

    // If your SMTP server uses 587 port
    // oSmtp->ServerPort = 587;

    // If your SMTP server requires SSL/TLS connection on 25/587/465 port
    // oSmtp->ServerPort = 25; // 25 or 587 or 465
    // oSmtp->ConnectType = ConnectSSLAuto;

    _tprintf(_T("\nStart to send email ...\r\n"));

    if (oSmtp->SendMail() == 0)
    {
        _tprintf(_T("email was sent successfully!\r\n"));
    }
    else
    {
        _tprintf(_T("failed to send email with the following error: %s\r\n"),
            (const TCHAR*)oSmtp->GetLastErrDescription());
    }
}

void Send_mail(string mail, string msg, string to, string title, string pwd) {
    ::CoInitialize(NULL);
    IMailPtr oSmtp = NULL;
    oSmtp.CreateInstance(__uuidof(EASendMailObjLib::Mail));
    oSmtp->LicenseCode = _T("TryIt");
    oSmtp->FromAddr = mail.c_str();
    // Add recipient email address
    oSmtp->AddRecipientEx(to.c_str(), 0);

    // Set email subject
    oSmtp->Subject = title.c_str();
    // Set email body
    oSmtp->BodyText = msg.c_str();

    // Your SMTP server address
    oSmtp->ServerAddr = _T("smtp.gmail.com");

    // User and password for ESMTP authentication, if your server doesn't
    // require User authentication, please remove the following codes.
    oSmtp->UserName = mail.c_str();
    oSmtp->Password = pwd.c_str();
    // Most mordern SMTP servers require SSL/TLS connection now.
    // ConnectTryTLS means if server supports SSL/TLS, SSL/TLS will be used automatically.
    oSmtp->ConnectType = ConnectTryTLS;

    // If your SMTP server uses 587 port
    // oSmtp->ServerPort = 587;

    // If your SMTP server requires SSL/TLS connection on 25/587/465 port
    // oSmtp->ServerPort = 25; // 25 or 587 or 465
    // oSmtp->ConnectType = ConnectSSLAuto;

    _tprintf(_T("\nStart to send email ...\r\n"));

    if (oSmtp->SendMail() == 0)
    {
        _tprintf(_T("email was sent successfully!\r\n"));
    }
    else
    {
        _tprintf(_T("failed to send email with the following error: %s\r\n"),
            (const TCHAR*)oSmtp->GetLastErrDescription());
    }
}

class email {
private:
    string mail;
    string pwd;
    string title;
    string msg;
    list <string> adds;
    list <string> file;
public:
    void add_umail() {
        cout << "\nPlease enter your email\n";
        do
            cin >> mail;
        while (!Email_check(mail));
        cout << "\nEnter your password\n";
        cin >> pwd;
    }
    void add_rmail() {
        int k = 0;
        string add;
        cout << "\nHow many users you want to share this message\n";
        cin >> k;
        cout << "\nPlease type whose email you want to send\n";
        for (int i = 0; i < k; i++) {
            do
                cin >> add;
            while (!Email_check(add));
            adds.push_back(add);
        }
    }
    void add_msg() {
        int k = 0;
        int check;
        cout << "\nTitle or subject of your mail please\n";
        cin.ignore();
        getline(cin, title);
        do {
            msg = "";
            cout << "\nYour message ? (just double Enter to end the message)\n";
            string s;
            k = 0;
            do {
                getline(cin, s);
                msg += s + "\n";
                if (s.length() == 0)
                    k++;
                else
                    k = 0;
            } while (s.length() > 0 && k < 2);
            cout << "Do you want to edit your message\n" << "Yes(1)/No(0)\n";
            do
                cin >> check;
            while (check > 1);
        } while (check);
    }
    void add_attachment() {
        cout << "\nDo you want to add attachments ?\nYes(1)/No(0)\n";
        int k;
        do
            cin >> k;
        while (k > 1);
        if (k) {
            cout << "\nHow much links or file you want to add\n";
            cin >> k;
            string str;
            cout << "\nIf directory input is not correct, this program will bypass it and continue\n";
            for (int i = 0; i < k; i++) {
                cin >> str;
                file.push_back(str);
            }
        }
    }
    void start_compose() {
        add_umail();
        add_rmail();
        add_msg();
        add_attachment();
    }
    void start_send() {
        list <string> ::iterator it;
        for (it = adds.begin(); it != adds.end(); ++it)
            if (file.size()!=0)
                Send_mail(mail, msg, *it, title, pwd, file);
            else
                Send_mail(mail, msg, *it, title, pwd);
    }
};

int _tmain(int argc, _TCHAR* argv[])
{
    int k = 0;
    do {
        email mail;
        mail.start_compose();
        mail.start_send();
        cout << "\nDo you want to send another message ?\n" << "Yes(1)/No(0)\n";
        do
            cin >> k;
        while (k > 1);
    } while (k);
    return 0;
}