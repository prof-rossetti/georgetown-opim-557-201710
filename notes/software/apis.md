# Software APIs Overview

A software application's **Application Programming Interface (API)** provides functionality and instructions sufficient to allow other programs to interface with it.

It is not uncommon for a system to also use its own public API to perform its own desired functionality.

Most of todays popular APIs are **web services** which accept HTTP requests at specified URLs and return responses to fulfill those requests. Here are some example APIs and API providers:

 + [New York Times APIs](http://developer.nytimes.com/docs)
 + [Google APIs](https://developers.google.com/apis-explorer/#p/)
 + [Twitter APIs](https://dev.twitter.com/rest/public)
 + [Facebook Social Graph API](https://developers.facebook.com/docs/graph-api)
 + [Instagram API](https://instagram.com/developer/endpoints/)
 + [Foursquare API](https://developer.foursquare.com/docs/)
 + [GitHub API](https://developer.github.com/v3/)
 + [Yelp API](https://www.yelp.com/developers/documentation/v2/overview)
 + [Flickr API](https://www.flickr.com/services/api/)
 + [Getty Images API](http://developers.gettyimages.com/en/)
 + [US Federal Elections Commission API](https://api.open.fec.gov/developers)
 + [Alpha Vantage (Stock Market) API](https://www.alphavantage.co/documentation/)

### Authentication

Many web services require developers to first register to obtain valid credentials in the form of an **API Key** (i.e. a secret token string) and subsequently authenticate by providing the key alongside each API request.

### Response Formats

The most common format for API response data is `JSON`, but some APIs alternatively or additionally provide response data in `XML` or `CSV` format.

Example CSV:

```csv
city,name,league
New York,Mets,Major
New York,Yankees,Major
Boston,Red Sox,Major
Washington,Nationals,Major
New Haven,Ravens,Minor
```

Example JSON:

```js
[
  {"city": "New York", "name": "Yankees", "league":"major"},
  {"city": "New York", "name": "Mets", "league":"major"},
  {"city": "Boston", "name": "Red Sox", "league":"major"},
  {"city": "New Haven", "name": "Ravens", "league":"minor"}
]
```

Example XML:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<teams>
  <team>
    <city>New York</city>
    <league>major</league>
    <name>Yankees</name>
  </team>
  <team>
    <city>New York</city>
    <league>major</league>
    <name>Mets</name>
  </team>
  <team>
    <city>Boston</city>
    <league>major</league>
    <name>Red Sox</name>
  </team>
  <team>
    <city>New Haven</city>
    <league>minor</league>
    <name>Ravens</name>
  </team>
</teams>
```

## URL Parameters

Many APIs allow you to specify URL parameters along with your request. These URL parameters are appended to the end of the base URL, starting with a single question mark (`?`) to denote the rest of the URL contains parameters. Then each parameter follows a convention where the name of the parameter is followed by an equal sign (`=`), which is followed by the desired parameter value. If there are multiple parameters, subsequent parameters after the first are separated by the ampersand character `&`.

Example request URL: https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=MSFT&outputsize=compact&apikey=demo. In this example, `https://www.alphavantage.co/query` is the base URL. And `function`, `symbol`, `outputsize`, and `apikey` are the names of URL parameters.
