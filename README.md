# xDrip Report
Программа формирует отчет, используя данные файла выгрузки **.CSV** из приложения **xDrip+**.

Результатом её работы будут 2 созданных файла: **PDF** и **Excel**.

ОС: **Windows**

![image](https://github.com/AnandSamir/xDrip_insulin_report/assets/40866955/65211dbb-09bc-444c-b18c-1374180cb626)

На одном листе располагается информация за **3 дня**:
+ все комментарии о питании и кол-ве углеводов
+ когда сделан укол и доза
+ 2 отметки после инъекции (через 1 и 2 часа), исключая время после **21.00** вечера и до **7.00** утра
+ 1-я отметка, когда глюкоза выходит за пределы нормы, **ГК <= 3.9** и **ГК >= 15**

Присутствует возможность выбора формата для отображения количества углеводов.

В обоих файлах данные выровнены и готовы к печати.

### Полезные ресурсы

https://nightscout.github.io/ - Документация Nightscout.

https://github.com/zreptil/nightscout-reporter - Веб-приложение для создания PDF-документов на основе данных Nightscout.

https://xdrip.readthedocs.io/en/latest/use/interapp/ - Использование веб-сервисов в xDrip+

https://www.youtube.com/watch?v=VzdRIhULSx0 - Быстрая настройка NightScout в Heroku

https://diadim.com.ua/dexcom/dexcom-detailed-instructions/ - Настройка xDrip+ на телефоне
