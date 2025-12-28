#include "Logger.h"
#include <ctime>
#include <filesystem>
#include <iostream>

namespace fs = std::filesystem;

std::time_t file_time_to_time_t(fs::file_time_type fileTime)
{
    using namespace std::chrono;

    // 把 filesystem 的时钟时间点转换成 system_clock
    auto sctp = time_point_cast<system_clock::duration>(
        fileTime - fs::file_time_type::clock::now() + system_clock::now()
    );

    return system_clock::to_time_t(sctp);
}

Logger& Logger::Instance()
{
    static Logger instance;
    return instance;
}

void Logger::Init(const std::string& filename)
{
    fs::path logPath = filename;
    fs::path dir = logPath.parent_path();

    // 如果路径中有目录部分，则确保存在
    if (!dir.empty() && !fs::exists(dir))
    {
        try {
            fs::create_directories(dir);
        }
        catch (const std::exception& e) {
            // 输出错误信息（可选）
            std::cerr << "Failed to create log directory: " << e.what() << std::endl;
        }
    }

    std::lock_guard<std::mutex> lock(mutex_);
    log_filename_ = filename;
    RotateOldFileByModifyTime();   // 重命名旧日志

    // 如果已有打开的流，先关闭它再重新打开
    if (ofs_.is_open()) {
        ofs_.close();
    }

    ofs_.open(log_filename_, std::ios::out | std::ios::app);
    if (!ofs_.is_open()) {
        std::cerr << "Failed to open log file: " << log_filename_ << std::endl;
    } else {
        // 使流在每次插入后自动 flush，增加实时性保障
        ofs_ << std::unitbuf;
    }
}

void Logger::Log(const std::string& level, const std::string& msg)
{
    std::lock_guard<std::mutex> lock(mutex_);
    if (!ofs_.is_open()) return;

    auto now = std::time(nullptr);
    char buf[32];
    std::strftime(buf, sizeof(buf), "%Y-%m-%d %H:%M:%S", std::localtime(&now));

    ofs_ << "[" << buf << "] [" << level << "] " << msg << std::endl;

    // 额外保证写入被刷新到底层文件
    try {
        ofs_.flush();
    }
    catch (...) {
        // 忽略 flush 失败，但可选地输出错误以便诊断
        std::cerr << "Logger: flush failed\n";
    }
}

void Logger::RotateOldFileByModifyTime()
{
    if (!fs::exists(log_filename_))
        return;

    // 获取文件的最后修改时间
    auto ftime = fs::last_write_time(log_filename_);
    auto t = file_time_to_time_t(ftime);

    std::tm tm{};
#ifdef _WIN32
    localtime_s(&tm, &t);
#else
    localtime_r(&t, &tm);
#endif

    // 生成时间戳字符串
    std::ostringstream oss;
    oss << std::put_time(&tm, "_%Y%m%d_%H%M%S");

    // 构造新文件名
    std::string backup = log_filename_;
    auto dot = backup.find_last_of('.');
    if (dot != std::string::npos)
        backup.insert(dot, oss.str());
    else
        backup += oss.str();

    // 重命名
    fs::rename(log_filename_, backup);
}