<div align="center">

<img src="https://capsule-render.vercel.app/api?type=rect&color=0:1a1c2c,100:4a1942&height=150&section=header&text=EXCEL%20MERGE%20CORE&fontSize=60&animation=fadeIn&fontColor=ff79c6" width="100%" style="border-radius: 18px;" />

<br>

<a href="https://git.io/typing-svg">
  <img src="https://readme-typing-svg.herokuapp.com?font=Fira+Code&weight=600&size=18&duration=3000&pause=1000&color=FF79C6&center=true&vCenter=true&width=500&height=30&lines=>>+SCANNING+DIRECTORIES...;>>+EXTRACTING+XLSX+DATA...;>>+GENERATING+CLIENT+REPORT;>>+STATUS:+STABLE+LEGACY" alt="Typing SVG" />
</a>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.x-blue?style=flat-square&logo=python">
  <img src="https://img.shields.io/badge/UI-PySimpleGUI-orange?style=flat-square">
  <img src="https://img.shields.io/badge/Data-Pandas-150458?style=flat-square&logo=pandas">
</p>

</div>

---

## 🛰 Описание | Overview

**Excel Merge Core** — это автономная утилита для консолидации данных. Скрипт автоматизирует рутинный процесс сборки одного большого отчета из множества разрозненных таблиц клиентов. 

> [!NOTE]
> Идеальное решение для задач, где нужно быстро "склеить" выгрузки за день/месяц без ручного копирования.

---

## ⚙️ Функционал (Logic)



- 🔍 **Auto-Scan:** Программа сама находит все `.xlsx` файлы в папке.
- 🚫 **Self-Exclusion:** Интеллектуальный пропуск файла `report.xlsx` для предотвращения рекурсии.
- 🧬 **Pandas Integration:** Быстрое слияние через `pd.concat` (обработка тысяч строк за секунды).
- 🖥️ **GUI Interface:** Минималистичное окно управления — никакой работы через консоль.

---

## 🛠 Технологический стек

* **Pandas** — обработка и объединение датафреймов.
* **OpenPyXL** — запись данных в ячейки с сохранением структуры.
* **PySimpleGUI** — графическая оболочка.
* **Xlwings** — работа с экземплярами Excel.

---

## 📂 Структура данных

```text
📁 Project_Folder
├── 📄 main.py            # Логика автоматизации
├── 📊 report.xlsx        # Целевой файл (куда пойдут данные)
├── 📈 data_01.xlsx       # Источник данных 1
└── 📈 data_02.xlsx       # Источник данных 2
