const instrument_id = '%s';
const start_date = '%s';
const end_date = '%s';
const api_url = `https://api.investing.com/api/financialdata/historical/${instrument_id}?start-date=${start_date}&end-date=${end_date}&time-frame=Monthly&add-missing-rows=false`;
const res = await fetch(api_url,{
    "headers":{
        "accept": "application/json, text/plain, */*",
        "accept-encoding": "gzip, deflate, br",
        "accept-language": "zh-CN,zh;q=0.9",
        "domain-id": "www",
        "origin": "https://www.investing.com",
        "referer": "https://www.investing.com/",
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-site",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36"
    },
    "body": null,
    "method":"GET"
});

return await res.json()