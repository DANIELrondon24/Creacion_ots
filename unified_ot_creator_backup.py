"""
Unified OT Creator Script

This script integrates the logic from CREACION_OT_Z63.py and CREACION_OT_Z67.py,
migrating the automation to Playwright for improved reliability and performance.
It provides a unified entry point that selects the appropriate functional unit logic
based on the mobile identifier prefix (Z63- or Z67-).

Features:
- Playwright-based automation (headless mode)
- Robust error handling with retries
- Comprehensive logging
- Configuration management via environment variables
- Input validation and edge case handling
- Modular code structure with type hints
- Production-ready with timeouts and scalability considerations

Requirements:
- Python 3.8+
- pandas
- python-dotenv
- playwright

Install dependencies:
    pip install pandas python-dotenv playwright
    playwright install

Usage:
    Set environment variables or create a .env file with:
    MAXIMO_USERNAME=your_username
    MAXIMO_PASSWORD=your_password
    EXCEL_FILE_PATH=path/to/your/excel/file.xlsx
    WORKSHOP_TIME=3
    HEADLESS=true  # optional, default true
    TIMEOUT=30000  # optional, default 30000ms

    python unified_ot_creator.py
"""

import os
import logging
import time
from typing import List, Dict, Any, Optional
from pathlib import Path
from dataclasses import dataclass
from enum import Enum

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, Page, Browser, BrowserContext, TimeoutError as PlaywrightTimeoutError
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type


# Load environment variables
load_dotenv()

# Constants
MAXIMO_URL = "http://maximo.greenmovil.com.co/maximo/webclient/login/login.jsp?event=loadapp&value=wotrack&uisessionid=274&_tt=qtoatqki6mn877iuj1j4or8p3f"
DEFAULT_TIMEOUT = 30000  # 30 seconds
RETRY_ATTEMPTS = 3
RETRY_WAIT_MIN = 1
RETRY_WAIT_MAX = 10


# Type definitions
class FunctionalUnit(Enum):
    Z63 = "Z63"
    Z67 = "Z67"

@dataclass
class Config:
    username: str
    password: str
    excel_file_path: str
    workshop_time: str
    supervisor_id: str
    headless: bool = True
    timeout: int = DEFAULT_TIMEOUT

    @classmethod
    def from_env(cls) -> 'Config':
        """Create Config from environment variables."""
        return cls(
            username=os.getenv('MAXIMO_USERNAME', ''),
            password=os.getenv('MAXIMO_PASSWORD', ''),
            excel_file_path=os.getenv('EXCEL_FILE_PATH', ''),
            workshop_time=os.getenv('WORKSHOP_TIME', '3'),
            supervisor_id=os.getenv('SUPERVISOR_ID', ''),
            headless=os.getenv('HEADLESS', 'true').lower() == 'true',
            timeout=int(os.getenv('TIMEOUT', str(DEFAULT_TIMEOUT)))
        )

# Word replacement dictionaries
TYPE_REPLACEMENTS = {
    "HALL": "MCO"
}

DESCRIPTION_REPLACEMENTS = {
    "GARANTIA": "GARA",
    "MTTO PREVENTIVO": "SISTEMATICA",
    "DAÑO OPERACIONAL": "DANO OPERACIONAL",
    "FALLA TECNICA": "FALLA TECNICA",
    "HALLAZGO ENTE GESTOR": "HALLAZGO ENTE GESTOR"
}

# Setup logging
def setup_logging() -> logging.Logger:
    """Setup logging configuration."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('ot_creator.log'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

logger = setup_logging()

class MaximoAutomationError(Exception):
    """Custom exception for Maximo automation errors."""
    pass

class MaximoAutomator:
    """Handles Playwright-based automation for Maximo OT creation."""

    def __init__(self, config: Config):
        self.config = config
        self.playwright = None
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None

    def __enter__(self):
        """Context manager entry."""
        self.start_browser()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.close_browser()

    def start_browser(self):
        """Start Playwright browser."""
        try:
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(
                headless=self.config.headless,
                args=['--ignore-certificate-errors', '--allow-insecure-localhost']
            )
            self.context = self.browser.new_context(viewport={'width': 2100, 'height': 1100})
            self.page = self.context.new_page()
            logger.info("Browser started successfully")
        except Exception as e:
            logger.error(f"Failed to start browser: {e}")
            raise MaximoAutomationError(f"Browser startup failed: {e}")

    def close_browser(self):
        """Close browser and cleanup."""
        if self.page:
            self.page.close()
        if self.context:
            self.context.close()
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()
        logger.info("Browser closed")

    @retry(
        stop=stop_after_attempt(RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=RETRY_WAIT_MIN, max=RETRY_WAIT_MAX),
        retry=retry_if_exception_type((PlaywrightTimeoutError, MaximoAutomationError))
    )
    def login(self):
        """Perform login to Maximo."""
        if not self.page:
            raise MaximoAutomationError("Browser not initialized")

        try:
            self.page.goto(MAXIMO_URL, timeout=self.config.timeout)
            self.page.fill("#username", self.config.username)
            self.page.fill("#password", self.config.password)
            self.page.click("#loginbutton")
            logger.info("Login successful")
        except Exception as e:
            logger.error(f"Login failed: {e}")
            raise MaximoAutomationError(f"Login failed: {e}")

    @retry(
        stop=stop_after_attempt(RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=RETRY_WAIT_MIN, max=RETRY_WAIT_MAX),
        retry=retry_if_exception_type((PlaywrightTimeoutError, MaximoAutomationError))
    )
    def select_site(self, unit: FunctionalUnit):
        """Select the functional unit site."""
        if not self.page:
            raise MaximoAutomationError("Browser not initialized")

        selector = {
            FunctionalUnit.Z63: "#toolactions_GRN_CHANGEDEFSITE3-tbb",
            FunctionalUnit.Z67: "#toolactions_GRN_CHANGEDEFSITE5-tbb"
        }.get(unit)

        if not selector:
            raise MaximoAutomationError(f"Unknown functional unit: {unit}")

        try:
            self.page.wait_for_selector(selector, timeout=self.config.timeout)
            self.page.click(selector)
            time.sleep(2)
            logger.info(f"Site selected: {unit.value}")
        except Exception as e:
            logger.error(f"Site selection failed for {unit.value}: {e}")
            raise MaximoAutomationError(f"Site selection failed: {e}")

    @retry(
        stop=stop_after_attempt(RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=RETRY_WAIT_MIN, max=RETRY_WAIT_MAX),
        retry=retry_if_exception_type((PlaywrightTimeoutError, MaximoAutomationError))
    )
    def return_to_list_view(self):
        """Click the 'Vista de lista' button to return to the list view."""
        if not self.page:
            raise MaximoAutomationError("Browser not initialized")

        try:
            logger.info("Returning to list view...")
            # Wait for and click the "Vista de lista" button
            self.page.get_by_role("button", name="Vista de lista").click()
            time.sleep(3)  # Wait for the view to load
            logger.info("Successfully returned to list view")
        except Exception as e:
            logger.error(f"Failed to return to list view: {e}")
            raise MaximoAutomationError(f"Return to list view failed: {e}")

    @retry(
        stop=stop_after_attempt(RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=RETRY_WAIT_MIN, max=RETRY_WAIT_MAX),
        retry=retry_if_exception_type((PlaywrightTimeoutError, MaximoAutomationError))
    )
    def create_new_ot(self):
        """Create a new OT."""
        if not self.page:
            raise MaximoAutomationError("Browser not initialized")

        try:
            self.page.wait_for_selector("#toolactions_INSERT-tbb", timeout=self.config.timeout * 1.5)
            self.page.click("#toolactions_INSERT-tbb")
            time.sleep(2)
            logger.info("New OT creation initiated")
        except Exception as e:
            logger.error(f"Failed to create new OT: {e}")
            self.page.reload()
            raise MaximoAutomationError(f"New OT creation failed: {e}")

    def fill_ot_form(self, row: pd.Series) -> str:
        """Fill the OT form with data from Excel row."""
        if not self.page:
            raise MaximoAutomationError("Browser not initialized")

        try:
            # Activity
            self.page.wait_for_selector("#mad3161b5-tb2", timeout=self.config.timeout)
            self.page.fill("#mad3161b5-tb2", row['Actividad'].title())
            time.sleep(0.5)
            self.page.press("#mad3161b5-tb2", "Tab")

            # License plate
            self.page.wait_for_selector("#m3b6a207f-tb", timeout=self.config.timeout)
            self.page.fill("#m3b6a207f-tb", str(row['Placa']))
            time.sleep(0.5)
            self.page.press("#m3b6a207f-tb", "Tab")
            time.sleep(1)

            # Type
            self.page.wait_for_selector("#me2096203-tb", timeout=self.config.timeout)
            tipo = str(row['Tipo'])
            for old, new in TYPE_REPLACEMENTS.items():
                tipo = tipo.replace(old, new)
            self.page.fill("#me2096203-tb", tipo)
            time.sleep(0.5)
            self.page.press("#me2096203-tb", "Tab")

            # Description Type
            self.page.wait_for_selector("#m78c05445-tb", timeout=self.config.timeout)
            des_tipo = str(row['Descripcion_Tipo'])
            for old, new in DESCRIPTION_REPLACEMENTS.items():
                des_tipo = des_tipo.replace(old, new)
            self.page.fill("#m78c05445-tb", des_tipo)
            time.sleep(0.5)
            self.page.press("#m78c05445-tb", "Tab")

            # Priority
            self.page.wait_for_selector("#m950e5295-tb", timeout=self.config.timeout)
            self.page.fill("#m950e5295-tb", self.config.workshop_time)
            time.sleep(0.5)
            self.page.press("#m950e5295-tb", "Tab")

            # Scheduled Start
            self.page.wait_for_selector("#m8b12679c-tb", timeout=self.config.timeout)
            fecha = row['Fecha']
            if isinstance(fecha, pd.Timestamp):
                fecha_str = fecha.strftime('%d/%m/%Y')
            else:
                fecha_str = str(fecha)
            self.page.fill("#m8b12679c-tb", fecha_str)
            time.sleep(0.5)
            self.page.press("#m8b12679c-tb", "Tab")
            time.sleep(1)

            # Supervisor
            self.page.wait_for_selector("#mb2eb834-tb", timeout=self.config.timeout)
            self.page.fill("#mb2eb834-tb", self.config.supervisor_id)
            time.sleep(0.5)
            self.page.press("#mb2eb834-tb", "Tab")
            time.sleep(1)

            # Extract and save OT
            time.sleep(0.5)
            ot_element = self.page.wait_for_selector(
                "xpath=/html/body/form/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/table/tbody/tr[3]/td/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td/div/table/tbody/tr/td/table/tbody/tr[1]/td[2]/input[1]",
                timeout=self.config.timeout
            )
            ot_value = ot_element.get_attribute("value")
            logger.info(f"OT created: {ot_value} for plate: {row['Placa']}")

            # Save OT
            save_button = self.page.wait_for_selector("#toolactions_SAVE-tbb", timeout=self.config.timeout)
            try:
                save_button.click()
            except Exception:
                # Try JavaScript click if regular click fails
                self.page.evaluate("document.querySelector('#toolactions_SAVE-tbb').click()")

            return ot_value

        except Exception as e:
            logger.error(f"Failed to fill OT form for plate {row['Placa']}: {e}")
            raise MaximoAutomationError(f"Form filling failed: {e}")

def validate_config(config: Config) -> None:
    """Validate configuration parameters."""
    if not config.username or not config.password:
        raise ValueError("Username and password are required")
    if not config.excel_file_path or not Path(config.excel_file_path).exists():
        raise ValueError(f"Excel file not found: {config.excel_file_path}")
    if not config.workshop_time or not config.supervisor_id:
        raise ValueError("Workshop time and supervisor ID are required")

def determine_functional_unit(df: pd.DataFrame) -> FunctionalUnit:
    """Determine functional unit from the first mobile identifier."""
    if df.empty:
        raise ValueError("Excel file is empty")

    first_mobile = str(df.iloc[0]['Movil']).upper()
    if first_mobile.startswith('Z63-'):
        return FunctionalUnit.Z63
    elif first_mobile.startswith('Z67-'):
        return FunctionalUnit.Z67
    else:
        raise ValueError(f"Invalid mobile identifier prefix: {first_mobile}. Must start with Z63- or Z67-")

def process_ot_creation(config: Config, unit: FunctionalUnit) -> List[str]:
    """Process OT creation for the specified functional unit."""
    df = pd.read_excel(config.excel_file_path)
    logger.info(f"Loaded {len(df)} rows from Excel file")
    logger.info(f"Processing for functional unit: {unit.value}")

    ots_created = []

    with MaximoAutomator(config) as automator:
        automator.login()
        automator.select_site(unit)

        for index, row in df.iterrows():
            try:
                automator.create_new_ot()
                ot = automator.fill_ot_form(row)
                if ot:
                    ots_created.append(ot)
                    time.sleep(3)
                else:
                    logger.warning(f"OT not saved for row {index}")
                    break
            except Exception as e:
                logger.error(f"Failed to process row {index}: {e}")
                break

    # Update Excel with OTs
    if len(ots_created) != len(df):
        logger.warning(f"Expected {len(df)} OTs, got {len(ots_created)}")
    df['OT'] = ots_created + [''] * (len(df) - len(ots_created))
    df.to_excel(config.excel_file_path, index=False)

    logger.info("OT creation process completed")
    return ots_created

def main():
    """Main entry point."""
    try:
        config = Config.from_env()
        validate_config(config)

        df = pd.read_excel(config.excel_file_path)
        unit = determine_functional_unit(df)

        logger.info("Starting unified OT creator")
        logger.info(f"Configuration: headless={config.headless}, timeout={config.timeout}ms")

        ots = process_ot_creation(config, unit)

        logger.info(f"Successfully created {len(ots)} OTs")
        print("Las Ordenes de trabajo han sido creadas, por favor verifica que estén correctas")

    except Exception as e:
        logger.error(f"Script execution failed: {e}")
        print(f"Error: {e}")
        raise

if __name__ == "__main__":
    main()