#pragma once

#include <fstream>
#include <string>
#include <mutex>
#include <iomanip>
#include <sstream>

class Logger
{
public:
    static Logger& Instance();

    // 初始化日志系统
    void Init(const std::string& filename);

    // 写日志
    void Log(const std::string& level, const std::string& msg);

private:
    Logger() = default;
    ~Logger() { if (ofs_.is_open()) ofs_.close(); }
    Logger(const Logger&) = delete;
    Logger& operator=(const Logger&) = delete;

    // 按文件修改时间重命名旧日志
    void RotateOldFileByModifyTime();

private:
    std::string log_filename_;
    std::ofstream ofs_;
    std::mutex mutex_;
};

// 简洁宏接口
#define LOG_INFO(msg)  Logger::Instance().Log("INFO", msg)
#define LOG_WARN(msg)  Logger::Instance().Log("WARN", msg)
#define LOG_ERROR(msg) Logger::Instance().Log("ERROR", msg)
