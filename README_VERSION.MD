Version			1.05
Release Date 	2023-08-31 

----------------------------------------------------------------------
2023-08-31	v1.05 	vshestakov
Добавлен функционал декодирования UTF-8, PrettyPrint (Настраивается в Excel)

* README.MD
+ README_VERSION.MD

SampleTestData.xlsx
  На вкладку Options в таблицу TransformationOptions
  Добавлен столбцы
  * TextFormat ('PP' for PrettyPrint, def - PLAIN")
  * Encoding ('WIN' to WIN1251, def - UTF8")

DataGenerator.bas
SampleTestDataLIST.xlsm
  Теперь умеет конвертировать XML в UTF-8, оформлять XML (pretty print)
  Читает найстройки из Excel (TransformationOptions TextFormat, Encoding)
  По-умолчанию Encoding - UTF-8, Форматирование - plain text
  Добавлен режим отладки с большей информацией (IS_DEBUG_MODE)

  
Обновлены файлы кейса-примера
  SampleTestData.testclass.ContractObjects.xml
  SampleTestData.testclass.Contracts.xml
  SampleTestData.TestData.xml
  
----------------------------------------------------------------------
2023-08-31	v1.04 	vshestakov
Модифицирован кейс-пример, теперь с кириллицей

DataGenerator.bas
  Добавлен режим отладки с большей информацией (IS_DEBUG_MODE)
  
Для файлов кейса-примера добавлены столбцы с кириллицей, для теста кодировки
  SampleTestData.xlsx
  TestData_to_DCB_CreateCntQuestsXML.xsl
  TestData_to_DCB_CreateContractXML.xsl
  SampleTestData.testclass.conquests.xml
  SampleTestData.testclass.contract.xml
	
----------------------------------------------------------------------
2023-08-31	v1.03
Добавлен функционал декодирования UTF-8, PrettyPrint. Настраивается в Excel

DCB_TestData.xlsx
  На вкладку Options в таблицу TransformationOptions
  Добавлен столбцы
  * TextFormat ('PP' for PrettyPrint, def - PLAIN")
  * Encoding ('WIN' to WIN1251, def - UTF8")

DataGenerator.bas
  Теперь умеет конвертировать XML в UTF-8
  Теперь умеет оформлять XML (pretty print)
  Читает найстройки из Excel (TransformationOptions TextFormat, Encoding)
  По-умолчанию Encoding - UTF-8, Форматирование - plain text

----------------------------------------------------------------------
2023-08-25	v1.0	vlad-shestakov@mail.ru
Базовая версия