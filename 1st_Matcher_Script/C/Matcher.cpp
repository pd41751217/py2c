#include <iostream>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>
#include <map>
#include <filesystem>
#include <thread>
#include <mutex>
#include <algorithm>
#include <cmath>
#include <future>
#include <regex>

#ifdef _WIN32
#include <windows.h>
#include <commdlg.h>
#include <shlobj.h>
#endif

// CSV/Excel-like data structure
using Row = std::vector<std::string>;
using DataFrame = std::vector<Row>;

// Column definitions
const std::vector<std::string> daily_cols = {
    "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", 
    "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK"
};

const std::vector<std::string> degree_cols = {
    "AQ", "AS", "AU", "AW", "AY", "BA", "BC", "BE", "BG", "BI", "BK"
};

class CSVReader {
public:
    static DataFrame readCSV(const std::string& filename) {
        DataFrame data;
        std::ifstream file(filename);
        if (!file.is_open()) {
            throw std::runtime_error("Cannot open file: " + filename);
        }
        
        std::string line;
        while (std::getline(file, line)) {
            Row row;
            std::stringstream ss(line);
            std::string cell;
            
            while (std::getline(ss, cell, ',')) {
                // Remove quotes and trim whitespace
                cell.erase(std::remove(cell.begin(), cell.end(), '"'), cell.end());
                cell.erase(0, cell.find_first_not_of(" \t"));
                cell.erase(cell.find_last_not_of(" \t") + 1);
                row.push_back(cell);
            }
            if (!row.empty()) {
                data.push_back(row);
            }
        }
        return data;
    }
    
    static void writeCSV(const DataFrame& data, const std::string& filename) {
        std::ofstream file(filename);
        if (!file.is_open()) {
            throw std::runtime_error("Cannot create file: " + filename);
        }
        
        for (const auto& row : data) {
            for (size_t i = 0; i < row.size(); ++i) {
                if (i > 0) file << ",";
                file << "\"" << row[i] << "\"";
            }
            file << "\n";
        }
    }
};

class DataProcessor {
private:
    std::mutex matches_mutex;
    
public:
    struct RowData {
        std::string player;
        std::map<std::string, std::string> data;
        std::string total;
        std::string winPercent;
    };
    
    RowData parseRowToDict(const Row& row) {
        RowData data;
        if (row.empty()) return data;
        
        data.player = row[0];
        
        // Parse pairs of key-value starting from index 1
        for (size_t i = 1; i < row.size() - 2; i += 2) {
            if (i + 1 < row.size()) {
                data.data[row[i]] = row[i + 1];
            }
        }
        
        if (row.size() >= 2) {
            data.total = row[row.size() - 2];
            data.winPercent = row[row.size() - 1];
        }
        
        return data;
    }
    
    bool degreeMatch(const std::string& daily_val, const std::string& hist_range) {
        try {
            std::regex range_regex(R"((\d+)-(\d+))");
            std::smatch matches;
            
            if (std::regex_match(hist_range, matches, range_regex)) {
                int low = std::stoi(matches[1]);
                int high = std::stoi(matches[2]);
                int daily_int = std::stoi(daily_val);
                return daily_int >= low && daily_int <= high;
            }
            return false;
        } catch (...) {
            return false;
        }
    }
    
    std::vector<Row> processChunk(const std::vector<std::pair<size_t, Row>>& chunk,
                                  const DataFrame& daily_df,
                                  const DataFrame& raw_daily_df) {
        std::vector<Row> matches;
        
        for (const auto& [idx, row] : chunk) {
            try {
                std::cout << "Processing row " << idx << "..." << std::endl;
                RowData hist_row = parseRowToDict(row);
                
                for (size_t i = 0; i < daily_df.size(); ++i) {
                    bool is_match = true;
                    const Row& daily_row = daily_df[i];
                    
                    // Check each column for matching
                    for (const auto& [col, hist_val] : hist_row.data) {
                        if (col == "WinPercent" || col == "Total") continue;
                        
                        // Find column index in daily data
                        auto col_it = std::find(daily_cols.begin(), daily_cols.end(), col);
                        if (col_it == daily_cols.end()) continue;
                        
                        size_t col_idx = std::distance(daily_cols.begin(), col_it) + 1; // +1 for Player column
                        if (col_idx >= daily_row.size()) continue;
                        
                        std::string daily_val = daily_row[col_idx];
                        
                        if (daily_val.empty() || hist_val.empty()) continue;
                        
                        // Check if this is a degree column
                        bool is_degree_col = std::find(degree_cols.begin(), degree_cols.end(), col) != degree_cols.end();
                        
                        if (is_degree_col) {
                            if (!degreeMatch(daily_val, hist_val)) {
                                is_match = false;
                                break;
                            }
                        } else {
                            if (daily_val != hist_val) {
                                is_match = false;
                                break;
                            }
                        }
                    }
                    
                    if (is_match) {
                        std::cout << "***** Found matching result for row " << idx << " *****" << std::endl;
                        Row matched_row = raw_daily_df[i];
                        matched_row.insert(matched_row.end(), row.begin(), row.end());
                        matches.push_back(matched_row);
                    }
                }
            } catch (const std::exception& e) {
                std::cout << "Error occurred in loop: " << e.what() << std::endl;
                continue;
            }
        }
        
        return matches;
    }
    
    DataFrame filterDailyData(const DataFrame& raw_daily_df) {
        DataFrame filtered_data;
        
        for (const auto& row : raw_daily_df) {
            Row filtered_row;
            
            // Extract column 0 (Player)
            if (!row.empty()) {
                filtered_row.push_back(row[0]);
            }
            
            // Extract columns 41-62 (indices 41-62)
            for (int i = 41; i <= 62 && i < static_cast<int>(row.size()); ++i) {
                filtered_row.push_back(row[i]);
            }
            
            if (!filtered_row.empty()) {
                filtered_data.push_back(filtered_row);
            }
        }
        
        return filtered_data;
    }
    
    void processFiles(const std::string& daily_file, const std::string& historical_folder) {
        try {
            // Check if files exist
            if (!std::filesystem::exists(daily_file) || !std::filesystem::exists(historical_folder)) {
                throw std::runtime_error("One or both files not found.");
            }
            
            // Read daily file
            DataFrame raw_daily_df = CSVReader::readCSV(daily_file);
            DataFrame daily_df = filterDailyData(raw_daily_df);
            
            std::vector<Row> all_matches;
            
            // Process each file in historical folder
            for (const auto& entry : std::filesystem::directory_iterator(historical_folder)) {
                if (entry.is_regular_file()) {
                    std::string file_path = entry.path().string();
                    std::string file_name = entry.path().filename().string();
                    
                    // Skip non-CSV files
                    std::string ext = file_path.substr(file_path.find_last_of('.'));
                    if (ext != ".csv" && ext != ".xlsx") {
                        continue;
                    }
                    
                    std::cout << "Processing file: " << file_name << std::endl;
                    
                    DataFrame raw_hist_df = CSVReader::readCSV(file_path);
                    
                    // Create chunks for parallel processing
                    std::vector<std::pair<size_t, Row>> all_rows;
                    for (size_t i = 0; i < raw_hist_df.size(); ++i) {
                        all_rows.emplace_back(i, raw_hist_df[i]);
                    }
                    
                    size_t num_threads = std::thread::hardware_concurrency();
                    if (num_threads == 0) num_threads = 4;
                    
                    size_t chunk_size = std::ceil(static_cast<double>(all_rows.size()) / num_threads);
                    std::vector<std::future<std::vector<Row>>> futures;
                    
                    // Process chunks in parallel
                    for (size_t i = 0; i < all_rows.size(); i += chunk_size) {
                        size_t end = std::min(i + chunk_size, all_rows.size());
                        std::vector<std::pair<size_t, Row>> chunk(all_rows.begin() + i, all_rows.begin() + end);
                        
                        futures.push_back(std::async(std::launch::async, [this, chunk, &daily_df, &raw_daily_df]() {
                            return processChunk(chunk, daily_df, raw_daily_df);
                        }));
                    }
                    
                    // Collect results
                    for (auto& future : futures) {
                        auto chunk_matches = future.get();
                        all_matches.insert(all_matches.end(), chunk_matches.begin(), chunk_matches.end());
                    }
                    
                    std::cout << "--------------- " << file_name << " Processed successfully ---------------" << std::endl;
                }
            }
            
            // Save results
            if (!all_matches.empty()) {
                // Remove empty rows
                all_matches.erase(
                    std::remove_if(all_matches.begin(), all_matches.end(),
                        [](const Row& row) { return row.empty(); }),
                    all_matches.end()
                );
                
                std::string output_path = daily_file.substr(0, daily_file.find_last_of('.')) + "_Matches.csv";
                CSVReader::writeCSV(all_matches, output_path);
                std::cout << "Processing finished. Results saved to: " << output_path << std::endl;
            } else {
                std::cout << "NO Matches found..." << std::endl;
            }
            
        } catch (const std::exception& e) {
            std::cerr << "Error occurred: " << e.what() << std::endl;
        }
    }
};

#ifdef _WIN32
std::string openFileDialog() {
    OPENFILENAME ofn;
    char szFile[260] = {0};
    
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = "CSV Files\0*.csv\0Excel Files\0*.xlsx\0All Files\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = NULL;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
    
    if (GetOpenFileName(&ofn)) {
        return std::string(szFile);
    }
    return "";
}

std::string openFolderDialog() {
    BROWSEINFO bi = {0};
    bi.lpszTitle = "Select Historical Folder";
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
    
    LPITEMIDLIST pidl = SHBrowseForFolder(&bi);
    if (pidl != nullptr) {
        char path[MAX_PATH];
        if (SHGetPathFromIDList(pidl, path)) {
            CoTaskMemFree(pidl);
            return std::string(path);
        }
        CoTaskMemFree(pidl);
    }
    return "";
}
#endif

int main() {
    std::cout << "Zmatcher C++ Version" << std::endl;
    std::cout << "===================" << std::endl;
    
    DataProcessor processor;
    
#ifdef _WIN32
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    
    std::cout << "Select daily file..." << std::endl;
    std::string daily_file = openFileDialog();
    if (daily_file.empty()) {
        std::cout << "No daily file selected. Exiting." << std::endl;
        CoUninitialize();
        return 1;
    }
    
    std::cout << "Select historical folder..." << std::endl;
    std::string historical_folder = openFolderDialog();
    if (historical_folder.empty()) {
        std::cout << "No historical folder selected. Exiting." << std::endl;
        CoUninitialize();
        return 1;
    }
    
    CoUninitialize();
#else
    // For non-Windows systems, use console input
    std::string daily_file, historical_folder;
    
    std::cout << "Enter daily file path: ";
    std::getline(std::cin, daily_file);
    
    std::cout << "Enter historical folder path: ";
    std::getline(std::cin, historical_folder);
#endif
    
    std::cout << "Daily file: " << daily_file << std::endl;
    std::cout << "Historical folder: " << historical_folder << std::endl;
    std::cout << "Processing..." << std::endl;
    
    processor.processFiles(daily_file, historical_folder);
    
    std::cout << "Press Enter to exit...";
    std::cin.get();
    
    return 0;
}