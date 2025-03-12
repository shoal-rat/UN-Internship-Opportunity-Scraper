# 🌐 UN Internship Opportunity Scraper



## ✨ Discover Global Opportunities with Ease

Are you searching for meaningful international internship experiences? The UN Internship Scraper automates the discovery of United Nations internship opportunities worldwide, helping you find the perfect position aligned with your career goals.

## 🚀 Features

- **Automated Discovery**: Scrapes all current UN internship listings in real-time
- **Global Organization**: Automatically categorizes opportunities by country
- **Application Tracking**: Filters out positions you've already applied for
- **Smart Geocoding**: Identifies countries from city names
- **Anti-Detection Measures**: Uses selenium-stealth to ensure reliable data collection

## 📊 What You Get

- **Comprehensive Excel Workbook**: Organized by country with separate sheets
- **Complete Details**: Job titles, IDs, departments, deadlines, and direct links
- **Geocoded Information**: City and country data for each position
- **Detailed Logging**: Track the scraping process and diagnose any issues

## 🛠️ Quick Setup

1. **Prerequisites**:
   ```bash
   pip install selenium selenium-stealth openpyxl geopy pycountry
   ```

2. **ChromeDriver**:
   - Download [ChromeDriver](https://chromedriver.chromium.org/downloads) matching your Chrome version
   - Place `chromedriver.exe` in the same directory as the script

3. **Application Tracking**:
   - Create an empty Excel file named `applied_intern.xlsx` in the script directory
   - Add a "Job ID" column to track positions you've applied for

4. **Run the Script**:
   ```bash
   python un_internship_scraper.py
   ```

## 🗂️ Output Structure

```
Project Directory
├── un_internship_scraper.py
├── chromedriver.exe
├── applied_intern.xlsx (your application tracking)
├── UN_Internships.xlsx (generated output)
└── logs/
    ├── scraping.log
    └── screenshots/
```

## 🌟 Why Use This Tool?

- **Save Time**: Automates hours of manual searching
- **Comprehensive Coverage**: Captures all available opportunities
- **Stay Organized**: Country-based organization helps target specific regions
- **Never Miss Deadlines**: All deadlines clearly listed in one place
- **Direct Access**: Click directly through to application pages

## 🔒 Privacy & Ethics

This tool is designed for personal use to streamline your internship search process. It respects the UN careers website's structure and implements reasonable delays to avoid overwhelming their servers. Always use responsibly.

## 🤝 Contributing

Contributions are welcome! Feel free to submit pull requests for new features, bug fixes, or documentation improvements.

## 📝 License

This project is available under the MIT License - helping students and young professionals discover career-building opportunities at the United Nations.

---

*"The best way to predict your future is to create it." - Abraham Lincoln*

Happy internship hunting!
