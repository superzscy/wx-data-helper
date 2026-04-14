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
    void OnPanelColumnsUpdated(wxCommandEvent& event);

    void StartExtract(wxCommandEvent& event);

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
    : wxFrame(nullptr, wxID_ANY, wxT("数据提取助手"))
{
    wxMenu* menuFile = new wxMenu;
    menuFile->Append(ID_Hello, wxT("问候"),
        wxT("测试菜单"));
    menuFile->AppendSeparator();
    menuFile->Append(wxID_EXIT);

    wxMenu* menuHelp = new wxMenu;
    menuHelp->Append(wxID_ABOUT);

    wxMenuBar* menuBar = new wxMenuBar;
    menuBar->Append(menuFile, wxT("文件"));
    menuBar->Append(menuHelp, wxT("帮助"));

    SetMenuBar(menuBar);

    CreateStatusBar();
    SetStatusText(wxT("就绪"));

    Bind(wxEVT_MENU, &MyFrame::OnHello, this, ID_Hello);
    Bind(wxEVT_MENU, &MyFrame::OnAbout, this, wxID_ABOUT);
    Bind(wxEVT_MENU, &MyFrame::OnExit, this, wxID_EXIT);
    Bind(wxEVT_EXCEL_PANEL_COLUMNS_UPDATED, &MyFrame::OnPanelColumnsUpdated, this);

    /////////////////////////////////////////////////////////

    wxBoxSizer* mainSizer = new wxBoxSizer(wxVERTICAL);

    // 使用说明：直接改这两个变量即可
    const wxString usageText = wxT("使用说明：\n1. 先加载表1与表2\n2. 选择匹配列和提取列\n3. 点击“开始提取”生成新文件");
    const int usageFontSize = 11;

    wxStaticText* usageTextBlock = new wxStaticText(this, wxID_ANY, usageText);
    wxFont usageFont = usageTextBlock->GetFont();
    usageFont.SetPointSize(usageFontSize);
    usageTextBlock->SetFont(usageFont);
    mainSizer->Add(usageTextBlock, 0, wxEXPAND | wxLEFT | wxRIGHT | wxTOP, 10);

    m_Panel1 = new ExcelInputPanel(this, wxT("表1："), false);
    mainSizer->Add(m_Panel1, 0, wxEXPAND | wxALL, 10);

    m_Panel2 = new ExcelInputPanel(this, wxT("表2："), true);
    mainSizer->Add(m_Panel2, 0, wxEXPAND | wxALL, 10);

    wxButton* btnStartMatch = new wxButton(this, wxID_ANY, wxT("开始提取"));
    btnStartMatch->Bind(wxEVT_BUTTON, &MyFrame::StartExtract, this);
    mainSizer->Add(btnStartMatch, 0, wxCENTER | wxALL, 10);

    this->SetSizer(mainSizer);
    this->Layout();
    this->Fit();
    wxSize size = this->GetSize();
    if (size.GetWidth() < 680) size.SetWidth(680);
    if (size.GetHeight() < 780) size.SetHeight(780);
    this->SetSize(size);
    this->SetMinSize(wxSize(680, 780));
}

void MyFrame::OnExit(wxCommandEvent& event)
{
    Close(true);
}

void MyFrame::OnAbout(wxCommandEvent& event)
{
    wxMessageBox("用于根据匹配列从表2提取字段并写回表1副本",
        "关于", wxOK | wxICON_INFORMATION);
}

void MyFrame::OnPanelColumnsUpdated(wxCommandEvent& event)
{
    if (!m_Panel1 || !m_Panel2)
    {
        return;
    }

    if (event.GetEventObject() == m_Panel1)
    {
        m_Panel2->ConfigureFilterBaseColumns(m_Panel1->GetDetectedColumnNames());
    }
}

void MyFrame::StartExtract(wxCommandEvent& event)
{
    if (!m_Panel1 || !m_Panel1->IsReady() || !m_Panel2 || !m_Panel2->IsReady())
    {
        wxMessageBox(wxT("请确保两个表都已加载，且参数有效。"), wxT("错误"), wxOK | wxICON_ERROR);
        return;
    }
    
    wxString errorMsg;
	bool bSucceed = false;
    
    do
    {
        wxBusyInfo info(wxT("正在提取数据，请稍候..."), this);

        bSucceed = m_Panel1->Parse();
        if (!bSucceed)
        {
            errorMsg = m_Panel1->GetErrorMsg();
            break;
        }
        m_Panel2->ConfigureFilterBaseColumns(m_Panel1->GetTitleNames());

        bSucceed = m_Panel2->Parse();
        if (!bSucceed)
        {
            errorMsg = m_Panel2->GetErrorMsg();
            break;
        }
        bSucceed = m_Panel1->Match(
            m_Panel2->GetContentRows(),
            m_Panel2->GetMatchColumnIndex(),
            m_Panel2->GetExtractColumnIndices(),
            m_Panel2->GetExtractColumnNames(),
            m_Panel2->ShouldFilter(),
            m_Panel2->GetFilterLeftColumnIndex(),
            m_Panel2->GetFilterRightColumnIndex());
        if (!bSucceed)
        {
            errorMsg = m_Panel1->GetErrorMsg();
        }
    } while (0);

    if (bSucceed)
    {
        wxMessageBox(wxString::Format(wxT("提取完成，结果已写入：\n%s"), m_Panel1->GetOutputFilePath()),
            wxT("信息"), wxOK | wxICON_INFORMATION);
    }
    else
    {
        wxMessageBox(errorMsg, wxT("错误"), wxOK | wxICON_ERROR);
    }
}

void MyFrame::OnHello(wxCommandEvent& event)
{
    LOG_INFO("来自wxWidgets的问候");
}
