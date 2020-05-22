# GAS Tinkoff Trades
![GAS Tinkoff Trades main image](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/main-image.jpg "GAS Tinkoff Trades main image")

Данный [Google Apps Script](https://developers.google.com/apps-script) предназначен для импорта сделок из Тинькофф Инвестиций прямо в Google таблицы, для последующего анализа. 

Я сделал этот скрипт для автоматизации ручного вбивания данных из приложения тинькофф, и надеюсь он окажется полезен кому-нибудь ещё :)

## Установка

* Создать или открыть документ Google Spreadsheets http://drive.google.com
* https://docs.google.com/spreadsheets/d/1cDjVeANfLXRECkYAPxkBjGEhZ61QdG_Pn9ap7I8J-VQ/edit?usp=sharing
* В меню "Tools" выбрать "Script Editor"
* Дать проекту имя, например `TinkoffTrades`
* Скопировать код из [Code.gs](https://raw.githubusercontent.com/ErhoSen/gas-tinkoff-trades/master/Code.gs)
* Получить [OpenApi-токен тинькофф](https://tinkoffcreditsystems.github.io/invest-openapi/auth/)
* Добавить свойство `OPENAPI_TOKEN` в разделе `File -> Project properties -> Script properties` равным токену, полученному выше. 
* Сохранить скрипт 💾

На этом всё. Теперь при работе с этим документом на всех листах будут доступны 2 новые функции `getPriceByTicker` и `getTrades`

## Функции

* `=getPriceByTicker(ticker, dummy)` - требует на вход [тикер](https://ru.wikipedia.org/wiki/%D0%A2%D0%B8%D0%BA%D0%B5%D1%80), и опциональный параметр `dummy`. Для автоматичекого обновления необходимо указать в качестве `dummy` ячейку `Z1`. 

* `=getTrades(ticker, from, to)` - требует на вход [тикер](https://ru.wikipedia.org/wiki/%D0%A2%D0%B8%D0%BA%D0%B5%D1%80), и опционально фильтрацию по времени. Параметры `from` и `to` являются строками и должны быть в [ISO 8601 формате](https://ru.wikipedia.org/wiki/ISO_8601)

## Особенности

* Скрипт резервирует ячейку `Z1` (самая правая ячейка первой строки), в которую вставляет случайное число на каждое изменении листа. Данная ячейка используется в функции `getPriceByTicker`, - она позволяет [автоматически обновлять](https://stackoverflow.com/a/27656313) текущую стоимость тикера при обновлении листа.

* Среди настроек скрипта есть `TRADING_START_AT` - дефолтная дата, начиная с которой фильтруются операции `getTrades`. По умолчанию это `Apr 01, 2020 10:00:00`, но данную константу можно в любой момент поменять в исходном коде.

## Пример использования 

```
=getPriceByTicker("V", Z1)  # Возвращает текущую цену акции Visa
=getPriceByTicker("FXMM", Z1)  # Возвращает текущую цену фонда казначейских облигаций США

=getTrades("V") 
# Вернёт все операции с акцией Visa, которые произошли начиная с TRADING_START_AT и по текущий момент.
=getTrades("V", "2020-05-01T00:00:00.000Z") 
# Вернёт все операции с акцией Visa, которые произошли начиная с 1 мая и по текущий моментs.
=getTrades("V", "2020-05-01T00:00:00.000Z", "2020-05-05T23:59:59.999Z") 
# Вернёт все операции с акцией Visa, которые произошли в период с 1 и по 5 мая.
```

## Пример работы

#### `=getTrades()`
![getTrades in action](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/get-trades-in-action.gif "getTrades in Action")

#### `=getPriceByTicker()`
![Get price by ticker in action](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/get-price-by-ticker.gif)
