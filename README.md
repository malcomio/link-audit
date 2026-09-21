# link-audit
Crawl a site and audit links for potential issues.

The script will create an Excel file with the following tabs:

| Name                      | Purpose                                                                                                                                                            |
|---------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Unsafe target_blank links | Links where `target` is `_blank`, but `rel` does not contain  `noopener` - see https://developer.mozilla.org/en-US/docs/Web/HTML/Reference/Attributes/rel/noopener |
| Failed page loads         | Internal links that result in an error, e.g. 404 not found                                                                                                         |
| Invalid links             | e.g. invalid protocol, email addresses without mailto                                                                                                              |
| Redirected links          | Internal links that result in a redirect                                                                                                                           |

## Running locally

1. Clone this repository
2. Install dependencies with `npm install`
3. Update `.env` with your configuration
   - `TEST_DOMAIN`: The domain to crawl
   - `BASIC_AUTH_USER`: Basic auth username (if required)
   - `BASIC_AUTH_PASSWORD`: Basic auth password (if required)
   - `CRAWL_MAX_PAGES`: The maximum number of pages to crawl
   - `CRAWL_DELAY_MS`: The delay between requests in milliseconds
 