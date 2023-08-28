# TestDataGenerator

Генератор тестовых данных использует Excel как точку ввода данных и конвертирует в подходящие XML.

Форк-версия проекта генерирует конкретные XML-тесты.

Доработано и основано на проекте **TestDataGenerator**

* Статья - [Генерация автоматических тестов: Excel, XML, XSLT, далее — везде](https://habr.com/ru/articles/312520/)

* GitHub - [TestDataGenerator](https://github.com/serhit/TestDataGenerator)

**Данные для тест-кейсов**

Собираются в Excel
<br>Гибкая настройка блоков за счет анализа свободных строк
 
![](resources/screenshot-excel.png)

**Гибкие настройки экспорта**

Произвольное количество тест-файлов за счет использования XSLT трансформации

![](resources/screenshot-options.png)

**Запуск пакета скриптов**

Генерация тестовых сценариев через макрос Visual Basic
</br>Возможность обработать несколько файлов за раз

![](resources/screenshot-macros.png)

## Примеры тест-файлов

**SampleTestData.TestData.xml** - промежуточный XML, результат преобразования из EXCEL

* [SampleTestData.TestData.xml](SampleTestData.TestData.xml)

**Сгенерированные XML Контракта и Объектов контракта**

* [SampleTestData.testclass.Contracts.xml](SampleTestData.testclass.Contracts.xml)

* [SampleTestData.testclass.ContractObjects.xml](SampleTestData.testclass.ContractObjects.xml)

## Как запустить

* Склонировать репозиторий в простой пусть (без русских букв и т.п.), например: C:\TestDataGenerator
* Открыть Excel (Тестировалось на Excel 2019)
* Открыть файл-теста [SampleTestData.xlsx](SampleTestData.xlsx)
	* На вкладке *"Options"* раздел ноастроек -  *"TransformationOptions"* содержит список *имен файлов-шаблонов XSLT*,соответствующее *выходное расширение файла*, имя выходного файла (опционально).
	* На остальных вкладках Excel *данные для тест-кейсов*, в нашем случае вкладка *"Case1"*. Их может быть несколько
* Открыть файл пакетной обработки [SampleTestDataLIST.xlsm](SampleTestDataLIST.xlsm)
<br>Здесь перечисляются файлы для обработки
* При открытии Excel появится предупреждение - Включить содержимое макросов
* В *ячейке "A1"* Указать путь для файла-теста, например на *C:\TestDataGenerator\SampleTestData.xlsx*
* Отобразить *Панель Разработчика* в Excel
	* /Панель меню Excel/Контекстное меню/Настройка ленты
	* /Список - Основные вкладки
	* Включить панель "Разработчик
* Запустить Макрос
	* /Панель Разработчика / Макросы
	* В окне выбрать функцию *"FileList_MakeDataExtractAndTransform"*
	* Нажать "Выполнить"
* Появится сообщение *"1 out of 1 tests has been processed"*
* В соответствии с указанными настройками в папке проекта появятся файлы:

	* [SampleTestData.TestData.xml](SampleTestData.TestData.xml) - промежуточный файл XML
	* [SampleTestData.testclass.Contracts.xml](SampleTestData.testclass.Contracts.xml) - Тестовый XML запрос *"Контракт"*
	* [SampleTestData.testclass.ContractObjects.xml](SampleTestData.testclass.ContractObjects.xml) - Тестовый XML запрос *"Объекты контракта"*


##  Статьи по XSLT и утилиты

* *[Справочник XSLT 1.0](https://xsltdev.ru/xslt/)*
* *[XML и XSLT в примерах для начинающих](http://citforum.ru/internet/xmlxslt/xmlxslt.shtml)*
* *[Online XSL Transformer (Тестирование XSLT преобразований](https://www.freeformatter.com/xsl-transformer.html)*