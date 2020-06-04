const scriptProperties = PropertiesService.getScriptProperties()
const CACHE = CacheService.getScriptCache()

const OPENAPI_TOKEN = scriptProperties.getProperty('OPENAPI_TOKEN')
const TRADING_START_AT = new Date('Apr 01, 2015 10:00:00')
const MILLIS_PER_DAY = 1000 * 60 * 60 * 24

let Currencies = new Map();
Currencies.set("USD", "USD000UTSTOM");
Currencies.set("EUR", "EUR_RUB__TOM");
Currencies.set("RUB", "MYRUB_TICKER");

function isoToDate(dateStr){
    // How to format date string so that google scripts recognizes it?
    // https://stackoverflow.com/a/17253060
    const str = dateStr.replace(/-/,'/').replace(/-/,'/').replace(/T/,' ').replace(/\+/,' \+').replace(/Z/,' +00')
    return new Date(str)
}

function getToday() {
    const now = new Date()
    let day = new Date(now + MILLIS_PER_DAY)
    //let day = new Date('Apr 01, 2020 10:00:00')
    return day.toISOString()
}

function onEdit(e) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    sheet.getRange('Z1').setValue(Math.random())
}

function makeDateISO(miles) {
    day = new Date(miles)
    day.setHours(0)
    day.setMinutes(0)
    day.setSeconds(0)
    day.setMilliseconds(0)
    return day.toISOString()
}

function getNextDate(day) {
    day = isoToDate(day)
    next = new Date(day.getTime() + MILLIS_PER_DAY)
    return next.toISOString()
}

function getPrevDate(day) {
    day = isoToDate(day)
    prev = new Date(day.getTime() - MILLIS_PER_DAY)
    return prev.toISOString()
}

class TinkoffClient {
    // Doc: https://tinkoffcreditsystems.github.io/invest-openapi/swagger-ui/
    // How to create a token: https://tinkoffcreditsystems.github.io/invest-openapi/auth/
    constructor(token) {
        this.token = token
        this.baseUrl = 'https://api-invest.tinkoff.ru/openapi/'
    }
  
    _makeApiCall(methodUrl) {
        const url = this.baseUrl + methodUrl
        Logger.log(`[API Call] ${url}`)
        const params = {'escaping': false, 'headers': {'accept': 'application/json', "Authorization": `Bearer ${this.token}`}}
        const response = UrlFetchApp.fetch(url, params)
        if (response.getResponseCode() == 200)
            return JSON.parse(response.getContentText())
    }

    getInstrumentByTicker(ticker) {
        const url = `market/search/by-ticker?ticker=${ticker}`
        const data = this._makeApiCall(url)
        return data.payload.instruments[0]
    }
    
    getInstrumentByFigi(figi) {
        const url = `market/search/by-figi?figi=${figi}`
        const data = this._makeApiCall(url)
        return data.payload
    }
  
    getOrderbookByFigi(figi) {
        const url = `market/orderbook?depth=1&figi=${figi}`
        const data = this._makeApiCall(url)
        return data.payload
    }
  
    getOperations(from, to, figi) {
        // Arguments `from` && `to` should be in ISO 8601 format
        const url = `operations?from=${from}&to=${to}&figi=${figi}`
        const data = this._makeApiCall(url)
        return data.payload.operations
    }
    
    getBrokerAccountId(brokerType) {
        const url = `user/accounts`
        const data = this._makeApiCall(url)
        const accounts = data.payload.accounts
        for (let q = 0; q < accounts.length; q++) {
            const {brokerAccountType, brokerAccountId} = accounts[q]
            if (brokerAccountType == brokerType) {
                return brokerAccountId
            }
        }
        return null
    }
    
    getAllOperations(from, to, brokerType) {
        // Arguments `from` && `to` should be in ISO 8601 format
        if (brokerType) {
            const brokerAccountId = this.getBrokerAccountId(brokerType)
            const url = `operations?from=${from}&to=${to}&brokerAccountId=${brokerAccountId}`
            const data = this._makeApiCall(url)
            return data.payload.operations
        } else {
            const url = `operations?from=${from}&to=${to}`
            const data = this._makeApiCall(url)
            return data.payload.operations
        }
    }
        
    getUserAccounts() {
        const url = `user/accounts`
        const data = this._makeApiCall(url)
        return data.payload.accounts
    }
    
    getPortfolio() {
        const url = `portfolio`
        const data = this._makeApiCall(url)
        return data.payload.positions
    }
    
    getPortfolioCurrencies(brokerType) {
        if (brokerType) {
            const brokerAccountId = this.getBrokerAccountId(brokerType)
            const url = `portfolio/currencies?brokerAccountId=${brokerAccountId}`
            const data = this._makeApiCall(url)
            return data.payload.currencies
        } else {
            const url = `portfolio/currencies`
            const data = this._makeApiCall(url)
            return data.payload.currencies
        }
    }
    
    getCandlesDay(figi, from, to) {
        const url = `market/candles?figi=${figi}&from=${from}&to=${to}&interval=day`
        const data = this._makeApiCall(url)
        return data.payload.candles
    }
    
}
    
const tinkoffClient = new TinkoffClient(OPENAPI_TOKEN)

function getNameByTicker(ticker) {
    if (!ticker) {
        return null
    }
    const cached = CACHE.get(`${ticker}_name`)
    if (cached != null) 
        return cached
    const {name} = tinkoffClient.getInstrumentByTicker(ticker)
    CACHE.put(`${ticker}_name`, name)
    return name
}
    
function _getFigiByTicker(ticker) {
    const cached = CACHE.get(ticker)
    if (cached != null) 
        return cached
    const {figi} = tinkoffClient.getInstrumentByTicker(ticker)
    CACHE.put(ticker, figi)
    return figi
}

function _getTickerByFigi(figi) {
    const cached = CACHE.get(figi)
    if (cached != null) 
        return cached
    const {ticker} = tinkoffClient.getInstrumentByFigi(figi)
    CACHE.put(figi, ticker)
    return ticker
}
    
function getPriceByTicker(ticker, dummy) {
    // dummy attribute uses for auto-refreshing the value each time the sheet is updating.
    // see https://stackoverflow.com/a/27656313
    if (!ticker) {
        return null
    }
    if (ticker == "MYRUB_TICKER" || ticker == "RUB") {
        return 1
    }
    if (Currencies.has(ticker)) {
        ticker = Currencies.get(ticker)
    }
    const figi = _getFigiByTicker(ticker)
    const {lastPrice} = tinkoffClient.getOrderbookByFigi(figi)
    return lastPrice
}

function getCurrencyPrice(name, day) {
    if (name == 'RUB') {
        return 1
    }
    if (!day) {
        let day_to = getToday()  
        day = getPrevDate(day_to)
    }
    let to = getNextDate(day)
    ticker = Currencies.get(name)
    figi = _getFigiByTicker(ticker)
    let data = tinkoffClient.getCandlesDay(figi, day, to)[0]
    while (data == undefined) {
        day = to
        to = getNextDate(day)
        data = tinkoffClient.getCandlesDay(figi, day, to)[0]
    }
    let price = data.o
    return price
}

function _calculateTrades(trades) {
  let totalSum = 0
  let totalQuantity = 0
  for (let j in trades) {
    const {quantity, price} = trades[j]
    totalQuantity += quantity
    totalSum += quantity * price
  }
  const weigthedPrice = totalSum / totalQuantity
  return [totalQuantity, totalSum, weigthedPrice]
}
    
function getTradesByTicker(ticker, from, to) {
  if (Currencies.has(ticker)) {
      ticker = Currencies.get(ticker)
  }
  const figi = _getFigiByTicker(ticker)
  if (!from) {
    from = TRADING_START_AT.toISOString()
  }
  if (!to) {
    to = getToday()
  }
  const operations = tinkoffClient.getOperations(from, to, figi)
  
  const values = [
    ["ID", "Date", "Operation", "Ticker", "Quantity", "Price", "Currency", "SUM", "Commission"], 
  ]
  for (let i=operations.length-1; i>=0; i--) {
    const {operationType, status, trades, id, date, currency, commission} = operations[i]
    if (operationType == "BrokerCommission" || status == "Decline") 
      continue
    let [totalQuantity, totalSum, weigthedPrice] = _calculateTrades(trades) // calculate weighted values
    if (operationType == "Buy") {  // inverse values in a way, that it will be easier to work with
      totalQuantity = -totalQuantity
      totalSum = -totalSum
    }
    values.push([
      id, isoToDate(date), operationType, ticker, totalQuantity, weigthedPrice, currency, totalSum, commission.value
    ])
  }
  return values
}

function getPortfolioCurrencies(brokerType) {
    const currencies = tinkoffClient.getPortfolioCurrencies(brokerType)
    let broker = `BROKER`
    if (brokerType && brokerType == `TinkoffIis`) {
        broker = `IIS`
    }
    
    
    let rates = new Map();
    rates.set("USD", getPriceByTicker(Currencies.get("USD")));
    rates.set("EUR", getPriceByTicker(Currencies.get("EUR")));
    rates.set("RUB", 1);
    
    const values = [
        ["BrokerAccount", "Currency", "Balance", "Balance in RUB"]
    ]
    for (let q = 0; q < currencies.length; q++) {
        const {currency, balance} = currencies[q]
        values.push(
            [broker, currency, balance, balance * rates.get(currency)]
        )
    }
    return values
}

function getAllPortfolioCurrencies() {
    const values = getPortfolioCurrencies("Tinkoff")
    const valuesIis = getPortfolioCurrencies("TinkoffIis")
    valuesIis.shift()
    return values.concat(valuesIis)
}

function getTrades(from, to, brokerType) {
    if (!from) {
        from = TRADING_START_AT.toISOString()
    }
    if (!to) {
        to = getToday()
    }
    let broker = `BROKER`
    if (brokerType && brokerType == `TinkoffIis`) {
        broker = `IIS`
    }
    const operations = tinkoffClient.getAllOperations(from, to, brokerType)
    
    const values = [
        ["Ticker", "Open", "Close", "Open date", "Close date", "Days", "Result in %", "BrokerAccount", "Quantity", "Sum", "Commission", "Result", "Currency", "Result in RUB"], 
    ]
    
    let portf = new Map();
    
    for (let q = operations.length - 1; q >= 0; q--) {
        const {status, commission, currency, trades, figi, date, operationType, instrumentType} = operations[q]
        
        if (status == "Decline" || operationType == "BrokerCommission" || instrumentType == "Currency" || instrumentType == "Bond") 
            continue
            
        if (operationType.indexOf(`Pay`) != -1)
            continue
            
        if (operationType == `MarginCommission` || operationType == `ServiceCommission` || operationType.indexOf(`Tax`) != -1)
            continue
            
        if (operationType == `Dividend` || operationType == `Coupon`)
            continue
            
        const [quantity, payment, price] = _calculateTrades(trades)
        
        const ticker = _getTickerByFigi(figi)
        
        let my_commission = 0
        if (commission) {
            my_commission = Math.abs(commission.value)
        }
        
        let my_date = isoToDate(date)
        
        let opType = `Buy`
        if (operationType.indexOf(`Sell`) != -1) {
            opType = `Sell`
        }
        let my_price = Math.abs(price)
        let my_payment = Math.abs(payment)
        
        let my_op = {
            'price': my_price, 
            'date': my_date, 
            'brokerAcc': broker, 
            'quantity': quantity, 
            'sum': my_payment, 
            'commission': my_commission, 
            'currency': currency, 
            'opType': opType,
            'dateISO': my_date.toISOString()
        }
        
        if (portf.has(ticker)) {
            for (let i = portf.get(ticker).length - 1; i >= 0; i--) {
                let op = portf.get(ticker)[i] 
                if (my_op.quantity * op.quantity == 0) {
                    continue
                }
                if (my_op.brokerAcc != op.brokerAcc) {
                    continue
                }
                if (my_op.opType != op.opType) {
                    let isSell = 1
                    if (op.opType == `Sell`) {
                        isSell *= -1
                    }
                    const qu = Math.min(my_op.quantity, op.quantity)
                    const days = Math.round((my_op.date - op.date) / MILLIS_PER_DAY)
                    const res_perc = Math.round((my_op.price / op.price - 1) * 100 * isSell * 100) / 100
                    const res = (my_op.price - op.price) * qu * isSell
                    const res_rub = res * getCurrencyPrice(my_op.currency, my_op.dateISO)
                    if (op.quantity >= my_op.quantity) {
                        let comis = my_op.commission + qu * op.commission / op.quantity
                        let val_op = [
                            ticker, op.price * isSell, my_op.price * isSell, op.date, my_op.date, days, res_perc, op.brokerAcc, qu, qu * op.price, comis, res, op.currency, res_rub
                        ]
                        op.quantity -= my_op.quantity
                        my_op.quantity = 0
                        values.push(val_op)
                    } else {
                        let comis = op.commission + qu * my_op.commission / my_op.quantity
                        let val_op = [
                            ticker, op.price * isSell, my_op.price * isSell, op.date, my_op.date, days, res_perc, op.brokerAcc, qu, qu * op.price, comis, res, op.currency, res_rub
                        ]
                        my_op.quantity -= op.quantity
                        op.quantity = 0
                        values.push(val_op)
                    }
                    if (op.quantity == 0) {
                        portf.get(ticker).pop()
                    }
                } else break
            }
            if (my_op.quantity != 0) {
                portf.get(ticker).push(my_op)
            }
        } else {
            portf.set(ticker, [
                my_op
            ])
        }
    }
    for (let ticker of portf.keys()) {
        for (let i = 0; i < portf.get(ticker).length; i++) {
            let op = portf.get(ticker)[i]
            let isSell = 1
            if (op.opType == `Sell`) {
                isSell *= -1
            }
            let val_op = [
                ticker, op.price * isSell, null, op.date, null, null, "OPEN", op.brokerAcc, op.quantity, op.sum, op.commission, null, op.currency, null
            ]
            values.push(val_op)
        }
    }
    return values
}

function getAllTrades(from, to) {
    const values = getTrades(from, to, "Tinkoff")
    const valuesIis = getTrades(from, to, "TinkoffIis")
    valuesIis.shift()
    return values.concat(valuesIis)
}

function getPays(from, to, brokerType) {
    if (!from) {
        from = TRADING_START_AT.toISOString()
    }
    if (!to) {
        to = getToday()
    }
    let broker = `BROKER`
    if (brokerType && brokerType == `TinkoffIis`) {
        broker = `IIS`
    }
  
    const values = [
        ["BrokerAccount", "Currency", "Payment", "Payment in RUB (on today)", "Date", "Operation Type"], 
    ]
    
    const operations = tinkoffClient.getAllOperations(from, to, brokerType)
    
    for (let i=operations.length-1; i>=0; i--) {
        const {status, currency, payment, date, operationType} = operations[i]
        
        if (status == "Decline") 
            continue
            
        if (operationType.indexOf(`Pay`) == -1)
            continue
          
        let my_date = isoToDate(date)
        let my_payment = payment * getCurrencyPrice(currency, makeDateISO(my_date))
    
        values.push([
            broker, currency, payment, my_payment, my_date, operationType
        ])
    }
    return values
}

function getAllPays(from, to) {
    const values = getPays(from, to, "Tinkoff")
    const valuesIis = getPays(from, to, "TinkoffIis")
    valuesIis.shift()
    return values.concat(valuesIis)    
}

function getTaxes(from, to, brokerType) {
    if (!from) {
        from = TRADING_START_AT.toISOString()
    }
    if (!to) {
        to = getToday()
    }
    let broker = `BROKER`
    if (brokerType && brokerType == `TinkoffIis`) {
        broker = `IIS`
    }
    
    const operations = tinkoffClient.getAllOperations(from, to, brokerType)
  
    const values = [
        ["BrokerAccount", "Currency", "Payment", "Payment in RUB (on today)", "Ticker", "Date", "Operation Type"], 
    ]
    for (let i=operations.length-1; i>=0; i--) {
        const {status, currency, payment, date, operationType, figi} = operations[i]
        
        if (status == "Decline") 
            continue
            
        if (operationType == `BrokerCommission`)
            continue
            
        if (operationType.indexOf(`Commission`) == -1 && operationType.indexOf(`Tax`) == -1)
            continue
          
        let ticker = null
        let my_date = isoToDate(date)
        let my_payment = payment * getCurrencyPrice(currency, my_date.toISOString())
        
        if (figi) {
            ticker = _getTickerByFigi(figi)
        }
    
        values.push([
            broker, currency, payment, my_payment, ticker, my_date, operationType
        ])
    }
    return values
}

function getAllTaxes(from, to) {
    const values = getTaxes(from, to, "Tinkoff")
    const valuesIis = getTaxes(from, to, "TinkoffIis")
    valuesIis.shift()
    return values.concat(valuesIis)    
}

function getDividends(from, to, brokerType) {
    if (!from) {
        from = TRADING_START_AT.toISOString()
    }
    if (!to) {
        to = getToday()
    }
    let broker = `BROKER`
    if (brokerType && brokerType == `TinkoffIis`) {
        broker = `IIS`
    }
    const operations = tinkoffClient.getAllOperations(from, to, brokerType)
  
    const values = [
        ["BrokerAccount", "Currency", "Payment", "Payment in RUB (on today)", "Ticker", "Date", "Operation Type"], 
    ]
    for (let i=operations.length-1; i>=0; i--) {
        const {status, currency, payment, date, operationType, figi} = operations[i]
        
        if (status == "Decline") 
            continue
            
        if (operationType != `Dividend` && operationType != `Coupon`)
            continue
          
        let ticker = null
        let my_date = isoToDate(date)
        let my_payment = payment * getCurrencyPrice(currency, my_date.toISOString())
        
        if (figi) {
            ticker = _getTickerByFigi(figi)
        }
    
        values.push([
            broker, currency, payment, my_payment, ticker, my_date, operationType
        ])
    }
    return values
}

function getAllDividends(from, to) {
    const values = getDividends(from, to, "Tinkoff")
    const valuesIis = getDividends(from, to, "TinkoffIis")
    valuesIis.shift()
    return values.concat(valuesIis)    
}
