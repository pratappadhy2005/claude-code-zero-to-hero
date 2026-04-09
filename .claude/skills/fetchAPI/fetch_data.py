"""Fetch CSV data from remote URLs and save with structured logging."""

import asyncio
import logging
from datetime import datetime
from pathlib import Path

import httpx

URLS = [
    "https://raw.githubusercontent.com/anshlambagit/AnshLambaYoutube/refs/heads/main/DBT_Masterclass/dim_customer.csv",
    "https://raw.githubusercontent.com/anshlambagit/AnshLambaYoutube/refs/heads/main/DBT_Masterclass/dim_store.csv",
    "https://raw.githubusercontent.com/anshlambagit/AnshLambaYoutube/refs/heads/main/DBT_Masterclass/dim_date.csv",
    "https://raw.githubusercontent.com/anshlambagit/AnshLambaYoutube/refs/heads/main/DBT_Masterclass/dim_product.csv",
    "https://raw.githubusercontent.com/anshlambagit/AnshLambaYoutube/refs/heads/main/DBT_Masterclass/fact_sales.csv",
    "https://raw.githubusercontent.com/anshlambagit/AnshLambaYoutube/refs/heads/main/DBT_Masterclass/fact_returns.csv",
]

BASE_DIR = Path(__file__).parent
TIMESTAMP = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

DATA_DIR = BASE_DIR / "data" / TIMESTAMP
LOG_DIR = BASE_DIR / "logs" / TIMESTAMP

DATA_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

# Configure logging
log_file = LOG_DIR / "fetchAPI.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)


async def fetch_url(client: httpx.AsyncClient, url: str) -> tuple[str, bool, str]:
    """Fetch a single URL and return (url, success, message)."""
    filename = url.split("/")[-1]
    logger.info("Fetching: %s", url)
    try:
        response = await client.get(url, timeout=30.0)
        response.raise_for_status()
        dest = DATA_DIR / filename
        dest.write_bytes(response.content)
        size = len(response.content)
        msg = f"SUCCESS — saved to {dest} ({size} bytes)"
        logger.info("%s → %s", filename, msg)
        return url, True, msg
    except httpx.HTTPStatusError as exc:
        msg = f"HTTP error {exc.response.status_code}"
        logger.error("%s → %s", filename, msg)
        return url, False, msg
    except Exception as exc:
        msg = f"Error: {exc}"
        logger.error("%s → %s", filename, msg)
        return url, False, msg


async def main() -> None:
    """Fetch all URLs concurrently and log results."""
    logger.info("=== fetchAPI run started — timestamp: %s ===", TIMESTAMP)
    logger.info("Data directory: %s", DATA_DIR)
    logger.info("Log directory:  %s", LOG_DIR)
    logger.info("Total URLs to fetch: %d", len(URLS))

    async with httpx.AsyncClient(follow_redirects=True) as client:
        results = await asyncio.gather(*[fetch_url(client, url) for url in URLS])

    successes = [r for r in results if r[1]]
    failures = [r for r in results if not r[1]]

    logger.info("=== Summary ===")
    logger.info("Successful: %d / %d", len(successes), len(URLS))
    if failures:
        logger.warning("Failed: %d / %d", len(failures), len(URLS))
        for url, _, msg in failures:
            logger.warning("  FAILED %s — %s", url, msg)
    else:
        logger.info("All fetches completed successfully.")
    logger.info("=== fetchAPI run complete ===")


if __name__ == "__main__":
    asyncio.run(main())
