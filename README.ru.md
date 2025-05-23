[English](README.md) | **Русский**
<p align="center">
  <img src="logo.svg" alt="" width="256">
</p>

---

# XLSX Генератор
Генерация таблиц Excel из простых текстовых данных. Дополнительный инструмент для AI

## Назначение скрипта
Этот скрипт преобразует текстовые файлы в форматированные Excel-файлы (XLSX) с автоматическим применением стилей:
- Автоматическое определение типов данных (числа, даты, текст)
- Цветовое выделение разных типов данных
- Форматирование границ и заголовков
- Автоподбор ширины столбцов

## Для пользователей

### Формат входного JSON-файла

Программа использует файл `xlsx_data.json` со следующей структурой:

```json
{
  "sheets": [
    {
      "name": "Sheet name",
      "columnWidths": [numbers],
      "rowHeights": [numbers],
      "data": [
        ["Cell1", "Cell2", ...],
        ["Cell1", 123, ...],
        ...
      ]
    }
  ]
}
```

### Элементы конфигурации

#### Основные параметры листа:
- `name` (необязательный) - название листа в книге
- `columnWidths` (необязательный) - массив ширин столбцов (в символах)
- `rowHeights` (необязательный) - массив высот строк (в пунктах)
- `data` (обязательный) - двумерный массив данных для листа

### Типы данных в ячейках

Программа автоматически определяет тип данных:
- **Строки**: `"text"`
- **Числа**: `123`, `45.67`
- **Булевы значения**: `true`, `false`
- **Формулы**: `"=SUM(A1:A10)"` (начинаются с =)
- **Даты**: строки в формате `"YYYY-MM-DD"`

### Автоматическое форматирование

Программа применяет стили автоматически:
1. **Верхняя строка** (заголовки):
   - Серый фон
   - Жирный текст
   - Выравнивание по центру
   - Перенос текста

2. **Первый столбец**:
   - Светло-серый фон
   - Жирный текст
   - Выравнивание по центру

3. **Числовые ячейки**:
   - Четные столбцы: бледно-голубые
   - Нечетные столбцы: бледно-синие

4. **Текстовые ячейки**:
   - Четные столбцы: бледно-зеленые
   - Нечетные столбцы: бледно-оранжевые

5. **Даты**:
   - Бледно-фиолетовый фон
   - Формат "YYYY-MM-DD"

### Примеры

#### Простая таблица:
```json
{
  "sheets": [
    {
      "name": "Report",
      "data": [
        ["Product", "Quantity", "Price"],
        ["Apples", 150, 25.50],
        ["Pears", 80, 32.75],
        ["Total", "=SUM(B2:B3)", "=SUM(C2:C3)"]
      ]
    }
  ]
}
```

#### Таблица с датами:
```json
{
  "sheets": [
    {
      "name": "Schedule",
      "columnWidths": [20, 15, 15],
      "data": [
        ["Event", "Date", "Days"],
        ["Project start", "2023-01-10", 0],
        ["First milestone", "2023-02-15", 36],
        ["Completion", "2023-05-20", 130]
      ]
    }
  ]
}
```

## Для разработчиков

### Общая структура

Программа состоит из одного основного класса `XlsxGenerator` с методами:
- `main()` - точка входа
- `generate()` - основной метод генерации
- Вспомогательные методы для создания стилей, листов и обработки данных

### Ключевые компоненты

#### Обработка входных данных
- `readJsonData()` - чтение и парсинг JSON
- Используется библиотека `org.json`

#### Генерация Excel
- Основана на Apache POI (`XSSFWorkbook`)
- Поддержка:
  - Различных типов данных
  - Формул
  - Стилей и форматирования
  - Автоматического определения размеров

#### Система стилей
- Стили создаются при инициализации
- Кэшируются в Map для повторного использования
- Автоматически применяются на основе:
  - Позиции ячейки
  - Типа данных
  - Четности столбца

### Обработка ошибок

- Базовая обработка исключений
- Режим отладки (`--debug`) для подробного вывода

### Расширяемость

#### Добавление новых стилей
1. Добавить цвет в массив `COLORS`
2. Создать новый стиль в `createDefaultStyles()`
3. Добавить логику применения в `applyCellStyle()`

#### Поддержка новых типов данных
Модифицировать метод `applyCellStyle()` для распознавания новых типов

### Важные зависимости

- **Apache POI** (v5.2.3+) - работа с Excel
- **org.json** (v20231013+) - обработка JSON

### Рекомендации по поддержке

1. **Тестирование**:
   - Проверять обработку всех типов данных
   - Тестировать граничные случаи (пустые данные, большие файлы)

2. **Логирование**:
   - Добавить более детальное логирование
   - Возможно интегрировать SLF4J

3. **Оптимизация**:
   - Для больших файлов рассмотреть потоковую обработку
   - Оптимизировать создание стилей

4. **Безопасность**:
   - Валидация входного JSON
   - Ограничение на размер файла

### Пример расширения

Для добавления поддержки гиперссылок:

```java
// В методе createSheet():
if (cellData instanceof String && ((String)cellData).startsWith("http")) {
    CreationHelper helper = workbook.getCreationHelper();
    Hyperlink link = helper.createHyperlink(HyperlinkType.URL);
    link.setAddress((String)cellData);
    cell.setHyperlink(link);
}
```
