// Start of wxWidgets "Hello World" Program
#include "ExcelInputPanel.h"
#include "Logger.h"
#include <filesystem>
#include <fstream>
#include <wx/busyinfo.h>
#include <wx/wx.h>
#include <wx/filedlg.h>
#include <wx/textdlg.h>
#include <xlnt/xlnt.hpp>  // Excel 读取库

class MyApp : public wxApp
{
public:
    bool OnInit() override;
};

wxIMPLEMENT_APP(MyApp);

class MyFrame : public wxFrame
{
public:
    MyFrame();

private:
    void OnHello(wxCommandEvent& event);
    void OnExit(wxCommandEvent& event);
    void OnAbout(wxCommandEvent& event);

    void StartMatch(wxCommandEvent& event);

    ExcelInputPanel* m_Panel1 = nullptr;
    ExcelInputPanel* m_Panel2 = nullptr;
};

enum
{
    ID_Hello = 1
};

bool MyApp::OnInit()
{
    Logger::Instance().Init("log/wx_data_helper.log");

    LOG_INFO("程序启动");

    MyFrame* frame = new MyFrame();
    frame->Show(true);
    return true;
}

MyFrame::MyFrame()
    : wxFrame(nullptr, wxID_ANY, wxT("数据小帮手"))
{
    wxMenu* menuFile = new wxMenu;
    menuFile->Append(ID_Hello, "&Hello...\tCtrl-H",
        "Help string shown in status bar for this menu item");
    menuFile->AppendSeparator();
    menuFile->Append(wxID_EXIT);

    wxMenu* menuHelp = new wxMenu;
    menuHelp->Append(wxID_ABOUT);

    wxMenuBar* menuBar = new wxMenuBar;
    menuBar->Append(menuFile, "&File");
    menuBar->Append(menuHelp, "&Help");

    SetMenuBar(menuBar);

    CreateStatusBar();
    SetStatusText("Welcome to wxWidgets!");

    Bind(wxEVT_MENU, &MyFrame::OnHello, this, ID_Hello);
    Bind(wxEVT_MENU, &MyFrame::OnAbout, this, wxID_ABOUT);
    Bind(wxEVT_MENU, &MyFrame::OnExit, this, wxID_EXIT);

    /////////////////////////////////////////////////////////

    wxBoxSizer* mainSizer = new wxBoxSizer(wxVERTICAL);

    m_Panel1 = new ExcelInputPanel(this, wxT("表1:"), false);
    mainSizer->Add(m_Panel1, 0, wxEXPAND | wxALL, 10);

    m_Panel2 = new ExcelInputPanel(this, wxT("表2:"), true);
    mainSizer->Add(m_Panel2, 0, wxEXPAND | wxALL, 10);

    wxButton* btnStartMatch = new wxButton(this, wxID_ANY, wxT("开始匹配"));
    btnStartMatch->Bind(wxEVT_BUTTON, &MyFrame::StartMatch, this);
    mainSizer->Add(btnStartMatch, 0, wxCENTER | wxALL, 10);

    this->SetSize(600, 500);
    this->SetSizer(mainSizer);
}

void MyFrame::OnExit(wxCommandEvent& event)
{
    Close(true);
}

void MyFrame::OnAbout(wxCommandEvent& event)
{
    wxMessageBox("This is a wxWidgets Hello World example",
        "About Hello World", wxOK | wxICON_INFORMATION);
}

void MyFrame::StartMatch(wxCommandEvent& event)
{
    if (!m_Panel1 || !m_Panel1->IsReady() || !m_Panel2 || !m_Panel2->IsReady())
    {
        wxMessageBox(wxT("请确保两个表格都已正确加载, 并且参数有效！"), wxT("错误"), wxOK | wxICON_ERROR);
        return;
    }
    
    wxString errorMsg;
	bool bSucceed = false;
    
    do
    {
        wxBusyInfo info(wxT("正在处理数据，请稍候..."), this);

        bSucceed = m_Panel1->Parse();
        if (!bSucceed)
        {
            errorMsg = m_Panel1->GetErrorMsg();
            break;
        }
        bSucceed = m_Panel2->Parse();
        if (!bSucceed)
        {
            errorMsg = m_Panel2->GetErrorMsg();
            break;
        }
        bSucceed = m_Panel1->Match(m_Panel2->GetContentRows(), m_Panel2->GetMatchColumnIndex(), m_Panel2->GetReturnColumnIndex());
        if (!bSucceed)
        {
            errorMsg = m_Panel1->GetErrorMsg();
        }
    } while (0);

    if (bSucceed)
    {
        wxMessageBox(wxString::Format(wxT("匹配完成, 结果已写入文件:\n%s"), m_Panel1->GetOutputFilePath()),
            wxT("信息"), wxOK | wxICON_INFORMATION);
    }
    else
    {
        wxMessageBox(errorMsg, wxT("错误"), wxOK | wxICON_ERROR);
    }
}

void MyFrame::OnHello(wxCommandEvent& event)
{
    LOG_INFO("Hello world from wxWidgets!");
}
