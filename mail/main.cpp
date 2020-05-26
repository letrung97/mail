#include "stdafx.h"
#include <tchar.h>
#include <Windows.h>
#include <string.h>
#include <iostream>
#include <regex>
#include "EASendMailObj.tlh"
#include <list>

using namespace EASendMailObjLib;

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

void Send_mail(string mail,string msg,string to,string title,string pwd) {
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

int _tmain(int argc, _TCHAR* argv[])
{

    string mail, pwd, title, msg;
    int k = 0;
    list <string> adds;
    string add;
    cout << "\nplease enter your email\n";
    do
        cin >> mail;
    while (!Email_check(mail));

    cout << "\nenter your password\n";
    cin >> pwd;

    cout << "\nPlease type whose email you want to send\n" << "how many users you want to share this message\n";
    cin >> k;
    for (int i = 0; i < k; i++) {
        do
            cin >> add;
        while (!Email_check(add));
        adds.push_back(add);
    }

    cout << "\nTitle or subject of your mail please\n";
    cin.ignore();
    getline(cin,title);

    cout << "\nyour message ? (just double Enter to end the message)\n";
    string s;
    k=0;
    do {
        getline(cin, s);
        msg += s + "\n";
        if (s.length() == 0)
            k++;
        else
            k = 0;
    } while (s.length() > 0 && k < 2);

    list <string> ::iterator it;
    for (it = adds.begin(); it != adds.end(); ++it)
        Send_mail(mail, msg, *it, title, pwd);

    return 0;
}