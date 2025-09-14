"""
Currency Handler for Material Management System

Handles currency formatting, conversion, and live exchange rates.
Integrates with BaseGenerator for consistent currency handling.

Author: Material Management System
Version: 1.0
"""

import requests
from decimal import Decimal, ROUND_HALF_UP
import os
from dotenv import load_dotenv

load_dotenv()
API_KEY = os.getenv("EXCHANGE_RATE_API_KEY")


class CurrencyHandler:
    """Currency handling with formatting, conversion, and live rates."""

    # Default static conversion rates (SAR to others)
    static_rates = {
        "SAR": 1.0,
        "USD": 0.27
    }

    def __init__(self, currency="SAR"):
        self.currency = currency.upper()
        if self.currency not in self.static_rates:
            raise ValueError(f"Unsupported currency: {self.currency}")

    def set_currency(self, currency_code):
        """Switch currency unit (SAR ↔ USD)."""
        currency_code = currency_code.upper()
        if currency_code not in self.static_rates:
            raise ValueError(f"Unsupported currency: {currency_code}")
        self.currency = currency_code

    def convert(self, amount, to_currency=None, live_rate=False):
        """
        Convert amount to target currency.
        
        Args:
            amount: float
            to_currency: str (optional)
            live_rate: bool (fetch rate from API if True)
        Returns:
            float
        """
        to_currency = to_currency.upper() if to_currency else self.currency

        if to_currency not in self.static_rates:
            raise ValueError(f"Unsupported currency: {to_currency}")

        rate = self.static_rates[to_currency]

        if live_rate and to_currency != "SAR":
            live = self.get_live_rate("SAR", to_currency)
            if live:
                rate = live

        converted = Decimal(float(amount) * rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return float(converted)

    def format(self, amount, currency=None, live_rate=False):
        """
        Format amount with currency symbol.
        Example: "USD 108.50"
        """
        currency = currency.upper() if currency else self.currency
        value = self.convert(amount, currency, live_rate=live_rate)
        return f"{currency} {value:,.2f}"

    def get_live_rate(self, from_currency, to_currency):
        """Fetch live exchange rate using ExchangeRate-API."""
        if not API_KEY:
            print("Warning: API key not set. Using static rate.")
            return None

        url = f"https://v6.exchangerate-api.com/v6/{API_KEY}/latest/{from_currency}"
        try:
            response = requests.get(url)
            data = response.json()
            if data.get("result") != "success":
                print("Error fetching live rate:", data.get("error-type"))
                return None
            rate = data["conversion_rates"].get(to_currency)
            if not rate:
                print(f"Currency {to_currency} not found in API response.")
            return rate
        except Exception as e:
            print("Error fetching live rate:", e)
            return None
