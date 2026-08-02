#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
============================================================================
🚀 FBS UNIT ECONOMICS PRO 2026 — ЯНДЕКС МАРКЕТ ВЕРСИЯ
============================================================================
Версия: 9.0.0 (Yandex Market Edition)
КЛЮЧЕВЫЕ ПРИНЦИПЫ:
1. НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ — все данные из API, AI, CSV или пользовательского ввода
2. ТОЛЬКО ЯНДЕКС МАРКЕТ — фокус на одном маркетплейсе для максимальной точности
3. ПОЛЬЗОВАТЕЛЬСКИЕ КАТЕГОРИИ — загрузка своих категорий с тарифами
4. ЖИВЫЕ ФОРМУЛЫ В EXCEL — при выгрузке сохраняются формулы для пересчёта
5. ГИБКИЙ МАППИНГ КОЛОНОК — обучение системы соответствию столбцов
6. ТРИ ЛИСТА В EXCEL: тарифы, расчёты, ABC/XYZ с дашбордами
7. НИКАКИХ СОКРАЩЕНИЙ — абсолютно полный код
8. ВЫБОР ИСТОЧНИКА ТАРИФОВ: API, AI, КОМБИНИРОВАННЫЙ
============================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import requests
import json
import re
import csv
import os
import base64
import hashlib
from pathlib import Path
import time
import logging
from typing import Dict, List, Any, Optional, Tuple, Union, Callable
from dataclasses import dataclass, field, asdict
from functools import wraps, lru_cache
import uuid
import math
import warnings
import io
import pickle
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
import traceback
from enum import Enum, auto
from abc import ABC, abstractmethod
import getpass
import sys
import gc

# ============================================================================
# БЛОК 0: БАЗОВАЯ КОНФИГУРАЦИЯ И НАСТРОЙКИ
# ============================================================================

APP_VERSION = "9.0.0"
APP_NAME = "🚀 FBS Юнит-экономика PRO 2026 — Яндекс Маркет"
APP_DESCRIPTION = "Профессиональный расчет юнит-экономики для FBS-модели на Яндекс Маркет"

# Настройка путей
BASE_DIR = Path(__file__).parent.resolve() if '__file__' in dir() else Path.cwd()
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
LOGS_DIR = BASE_DIR / "logs"
EXPORTS_DIR = BASE_DIR / "exports"
CONFIG_DIR = BASE_DIR / "config"
TEMP_DIR = BASE_DIR / "temp"
TARIFFS_CACHE_DIR = CACHE_DIR / "tariffs"
USER_CATEGORIES_DIR = DATA_DIR / "user_categories"
MAPPING_DIR = CONFIG_DIR / "mappings"          # Для сохранения маппингов колонок

for dir_path in [DATA_DIR, CACHE_DIR, LOGS_DIR, EXPORTS_DIR, CONFIG_DIR, TEMP_DIR,
                 TARIFFS_CACHE_DIR, USER_CATEGORIES_DIR, MAPPING_DIR]:
    try:
        dir_path.mkdir(exist_ok=True, parents=True)
    except Exception as e:
        print(f"⚠️ Не удалось создать директорию {dir_path}: {e}")

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    handlers=[
        logging.FileHandler(LOGS_DIR / "fbs_unit_economy.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('FBSEconomy')
logger.info(f"🚀 Запуск {APP_NAME} версии {APP_VERSION}")

# Попытка импорта дополнительных библиотек
try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
    logger.info("✅ Cryptography доступен")
except ImportError:
    CRYPTO_AVAILABLE = False
    Fernet = None
    logger.warning("⚠️ Cryptography не установлен. Шифрование будет отключено.")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle, numbers
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
    from openpyxl.chart import BarChart, Reference, PieChart, LineChart
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.hyperlink import Hyperlink
    from openpyxl.drawing.image import Image
    OPENPYXL_AVAILABLE = True
    logger.info("✅ OpenPyXL доступен")
except ImportError:
    OPENPYXL_AVAILABLE = False
    logger.warning("⚠️ OpenPyXL не установлен. Экспорт в Excel будет недоступен.")

try:
    import gspread
    from google.oauth2.service_account import Credentials
    from google.oauth2 import service_account
    GSPREAD_AVAILABLE = True
    logger.info("✅ GSpread доступен")
except ImportError:
    GSPREAD_AVAILABLE = False
    gspread = None
    logger.warning("⚠️ GSpread не установлен. Интеграция с Google Sheets будет недоступна.")

warnings.filterwarnings('ignore')


# ============================================================================
# БЛОК 1: ДЕКОРАТОРЫ И УТИЛИТЫ
# ============================================================================

def timing_decorator(func):
    """
    Декоратор для измерения времени выполнения функции.
    Логирует время выполнения, если оно превышает 1 секунду.
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        execution_time = end_time - start_time
        if execution_time > 1.0:
            logger.info(f"⏱️ {func.__name__} выполнена за {execution_time:.2f} сек")
        return result
    return wrapper

def retry_on_failure(max_retries: int = 3, delay: float = 1.0):
    """
    Декоратор для повторных попыток выполнения функции при ошибках.
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            last_exception = None
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    last_exception = e
                    logger.warning(f"⚠️ Попытка {attempt + 1}/{max_retries} для {func.__name__} не удалась: {e}")
                    if attempt < max_retries - 1:
                        time.sleep(delay * (attempt + 1))
            raise last_exception
        return wrapper
    return decorator

def memoize(func):
    """
    Декоратор для кэширования результатов функции.
    """
    cache = {}
    @wraps(func)
    def wrapper(*args, **kwargs):
        key = (args, frozenset(kwargs.items()))
        if key not in cache:
            cache[key] = func(*args, **kwargs)
        return cache[key]
    wrapper.cache_clear = cache.clear
    return wrapper

class ProgressTracker:
    """Класс для отслеживания прогресса выполнения длительных операций."""
    def __init__(self):
        self.progress = 0.0
        self.status = ""
        self.total = 0
        self.current = 0
        self.start_time = None
        self.estimated_time_remaining = 0
        self.history = []

    def start(self, total: int, status: str = ""):
        self.total = total
        self.current = 0
        self.progress = 0.0
        self.status = status
        self.start_time = time.time()
        self.history.append({
            'timestamp': datetime.now().isoformat(),
            'action': 'start',
            'total': total,
            'status': status
        })
        logger.info(f"📊 Начат процесс: {status} (всего: {total} шагов)")

    def update(self, current: int, status: str = ""):
        self.current = current
        self.total = max(self.total, current)
        self.progress = min(current / self.total, 1.0) if self.total > 0 else 0
        if status:
            self.status = status
        if self.start_time and self.progress > 0:
            elapsed = time.time() - self.start_time
            self.estimated_time_remaining = (elapsed / self.progress) * (1 - self.progress)
        if int(self.progress * 100) % 10 == 0 and self.progress > 0:
            logger.info(f"📊 Прогресс: {self.progress*100:.0f}% - {self.status}")

    def get_progress(self) -> float:
        return self.progress

    def get_status(self) -> str:
        return self.status

    def get_eta(self) -> float:
        return self.estimated_time_remaining

    def finish(self, status: str = "Завершено"):
        self.progress = 1.0
        self.status = status
        self.history.append({
            'timestamp': datetime.now().isoformat(),
            'action': 'finish',
            'total': self.total,
            'status': status,
            'duration_seconds': time.time() - self.start_time if self.start_time else 0
        })
        logger.info(f"✅ Процесс завершён: {status}")

class AuditLogger:
    """Класс для ведения аудиторского журнала всех действий пользователя."""
    def __init__(self):
        self.audit_file = LOGS_DIR / "audit.log"
        self._init_audit_file()

    def _init_audit_file(self):
        if not self.audit_file.exists():
            try:
                with open(self.audit_file, 'w', encoding='utf-8') as f:
                    f.write("timestamp,user,action,details\n")
                logger.info("✅ Файл аудита создан")
            except Exception as e:
                logger.error(f"❌ Ошибка создания файла аудита: {e}")

    def log(self, action: str, details: Dict[str, Any]):
        try:
            user = getpass.getuser()
            timestamp = datetime.now().isoformat()
            with open(self.audit_file, 'a', encoding='utf-8') as f:
                f.write(f"{timestamp},{user},{action},{json.dumps(details, ensure_ascii=False)}\n")
            logger.info(f"📝 Аудит: {user} - {action}")
        except Exception as e:
            logger.error(f"❌ Ошибка записи аудита: {e}")

    def get_logs(self, limit: int = 100) -> List[Dict]:
        logs = []
        try:
            if self.audit_file.exists():
                with open(self.audit_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()[1:]
                    for line in lines[-limit:]:
                        parts = line.strip().split(',', 3)
                        if len(parts) == 4:
                            logs.append({
                                'timestamp': parts[0],
                                'user': parts[1],
                                'action': parts[2],
                                'details': json.loads(parts[3]) if parts[3] else {}
                            })
        except Exception as e:
            logger.error(f"❌ Ошибка чтения аудита: {e}")
        return logs


# ============================================================================
# БЛОК 2: БЕЗОПАСНОЕ ХРАНЕНИЕ ДАННЫХ (ШИФРОВАНИЕ)
# ============================================================================

class SecureDataManager:
    """Класс для безопасного хранения данных с шифрованием."""
    def __init__(self):
        self.key_file = CONFIG_DIR / ".master_key"
        self.data_file = CONFIG_DIR / ".secure_data.enc"
        self._fernet = None
        self._init_encryption()
        self.audit_logger = AuditLogger()

    def _init_encryption(self):
        if not CRYPTO_AVAILABLE:
            logger.warning("⚠️ Cryptography не установлен. Данные не будут зашифрованы.")
            return
        try:
            if self.key_file.exists():
                key = self.key_file.read_bytes()
                logger.debug("🔑 Мастер-ключ загружен из файла")
            else:
                key = Fernet.generate_key()
                self.key_file.write_bytes(key)
                try:
                    os.chmod(self.key_file, 0o600)
                except OSError:
                    pass
                logger.info("🔑 Новый мастер-ключ создан")
            self._fernet = Fernet(key)
            logger.info("✅ Шифрование инициализировано успешно")
        except Exception as e:
            logger.error(f"❌ Ошибка инициализации шифрования: {e}")
            self._fernet = None

    def is_available(self) -> bool:
        return self._fernet is not None

    def save_data(self, data: Dict[str, Any]) -> bool:
        if not self._fernet:
            logger.warning("⚠️ Шифрование недоступно, данные не сохранены")
            return False
        try:
            json_data = json.dumps(data, ensure_ascii=False, indent=2)
            encrypted = self._fernet.encrypt(json_data.encode('utf-8'))
            self.data_file.write_bytes(encrypted)
            try:
                os.chmod(self.data_file, 0o600)
            except OSError:
                pass
            self.audit_logger.log('secure_data_save', {'size': len(encrypted)})
            logger.info("✅ Данные успешно зашифрованы и сохранены")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения зашифрованных данных: {e}")
            return False

    def load_data(self) -> Dict[str, Any]:
        if not self._fernet or not self.data_file.exists():
            return {}
        try:
            encrypted = self.data_file.read_bytes()
            decrypted = self._fernet.decrypt(encrypted)
            data = json.loads(decrypted.decode('utf-8'))
            logger.info("✅ Данные успешно расшифрованы и загружены")
            return data
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки зашифрованных данных: {e}")
            return {}

    def delete_data(self) -> bool:
        try:
            if self.data_file.exists():
                self.data_file.unlink()
                self.audit_logger.log('secure_data_delete', {})
                logger.info("🗑️ Зашифрованные данные удалены")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка удаления данных: {e}")
            return False

    def store_api_key(self, service: str, api_key: str) -> bool:
        data = self.load_data()
        if 'api_keys' not in data:
            data['api_keys'] = {}
        data['api_keys'][service] = api_key
        data['api_keys_updated'] = datetime.now().isoformat()
        success = self.save_data(data)
        if success:
            self.audit_logger.log('api_key_stored', {'service': service})
        return success

    def get_api_key(self, service: str) -> Optional[str]:
        data = self.load_data()
        return data.get('api_keys', {}).get(service)

    def get_all_api_keys(self) -> Dict[str, str]:
        data = self.load_data()
        return data.get('api_keys', {})

    def delete_api_key(self, service: str) -> bool:
        data = self.load_data()
        if 'api_keys' in data and service in data['api_keys']:
            del data['api_keys'][service]
            success = self.save_data(data)
            if success:
                self.audit_logger.log('api_key_deleted', {'service': service})
            return success
        return False

    def clear_all_keys(self) -> bool:
        data = self.load_data()
        data['api_keys'] = {}
        success = self.save_data(data)
        if success:
            self.audit_logger.log('all_api_keys_deleted', {})
        return success


# ============================================================================
# БЛОК 3: КЭШИРОВАНИЕ ДЛЯ ОПТИМИЗАЦИИ ПРОИЗВОДИТЕЛЬНОСТИ
# ============================================================================

class CacheManager:
    """Менеджер кэширования данных в памяти и на диске."""
    def __init__(self, max_memory_mb: int = 500, cache_ttl_seconds: int = 3600):
        self.cache_dir = CACHE_DIR
        self.max_memory_mb = max_memory_mb
        self.cache_ttl = cache_ttl_seconds
        self._memory_cache: Dict[str, Any] = {}
        self._cache_timestamps: Dict[str, float] = {}
        self._cache_sizes: Dict[str, int] = {}
        self._access_count: Dict[str, int] = {}
        self.tariffs_cache_dir = self.cache_dir / "tariffs"
        self.api_cache_dir = self.cache_dir / "api_responses"
        self.calc_cache_dir = self.cache_dir / "calculations"
        self.user_categories_cache_dir = self.cache_dir / "user_categories"
        for dir_path in [self.tariffs_cache_dir, self.api_cache_dir,
                        self.calc_cache_dir, self.user_categories_cache_dir]:
            dir_path.mkdir(exist_ok=True, parents=True)
        logger.info(f"✅ CacheManager инициализирован (TTL: {cache_ttl_seconds}с, "
                   f"Max: {max_memory_mb}МБ)")

    def _get_cache_key(self, *args, **kwargs) -> str:
        key_parts = [str(arg) for arg in args]
        key_parts.extend(f"{k}={v}" for k, v in sorted(kwargs.items()))
        key_string = "|".join(key_parts)
        return hashlib.md5(key_string.encode()).hexdigest()

    def _get_disk_cache_path(self, cache_type: str, key: str) -> Path:
        cache_dirs = {
            'tariffs': self.tariffs_cache_dir,
            'api': self.api_cache_dir,
            'calc': self.calc_cache_dir,
            'user_categories': self.user_categories_cache_dir
        }
        cache_dir = cache_dirs.get(cache_type, self.cache_dir)
        return cache_dir / f"{key}.cache"

    def get(self, cache_type: str, key: str) -> Optional[Any]:
        memory_key = f"{cache_type}:{key}"
        if memory_key in self._memory_cache:
            timestamp = self._cache_timestamps.get(memory_key, 0)
            if time.time() - timestamp < self.cache_ttl:
                self._access_count[memory_key] = self._access_count.get(memory_key, 0) + 1
                logger.debug(f"📦 Кэш попадание (память): {memory_key}")
                return self._memory_cache[memory_key]
            else:
                del self._memory_cache[memory_key]
                del self._cache_timestamps[memory_key]
                if memory_key in self._access_count:
                    del self._access_count[memory_key]
        cache_path = self._get_disk_cache_path(cache_type, key)
        if cache_path.exists():
            try:
                with open(cache_path, 'rb') as f:
                    cached_data = pickle.load(f)
                cached_time = cached_data.get('timestamp', 0)
                if time.time() - cached_time < self.cache_ttl:
                    logger.debug(f"💾 Кэш попадание (диск): {key}")
                    value = cached_data.get('data')
                    self._memory_cache[memory_key] = value
                    self._cache_timestamps[memory_key] = cached_time
                    self._access_count[memory_key] = 1
                    return value
                else:
                    cache_path.unlink()
                    logger.debug(f"🗑️ Удалён устаревший кэш: {cache_path}")
            except Exception as e:
                logger.debug(f"Ошибка чтения дискового кэша: {e}")
                try:
                    cache_path.unlink()
                except:
                    pass
        return None

    def set(self, cache_type: str, key: str, value: Any):
        memory_key = f"{cache_type}:{key}"
        current_time = time.time()
        self._memory_cache[memory_key] = value
        self._cache_timestamps[memory_key] = current_time
        self._access_count[memory_key] = 1
        try:
            cache_path = self._get_disk_cache_path(cache_type, key)
            cache_data = {
                'timestamp': current_time,
                'data': value,
                'cache_type': cache_type,
                'key': key,
                'created_at': datetime.now().isoformat()
            }
            with open(cache_path, 'wb') as f:
                pickle.dump(cache_data, f, protocol=pickle.HIGHEST_PROTOCOL)
            logger.debug(f"💾 Данные сохранены в кэш: {cache_path}")
        except Exception as e:
            logger.debug(f"Не удалось сохранить кэш на диск: {e}")
        self._cleanup_memory_cache()

    def _cleanup_memory_cache(self):
        current_time = time.time()
        expired_keys = []
        for key, timestamp in self._cache_timestamps.items():
            if current_time - timestamp > self.cache_ttl:
                expired_keys.append(key)
        for key in expired_keys:
            if key in self._memory_cache:
                del self._memory_cache[key]
            if key in self._cache_timestamps:
                del self._cache_timestamps[key]
            if key in self._access_count:
                del self._access_count[key]
        if expired_keys:
            logger.debug(f"🗑️ Очищено {len(expired_keys)} устаревших записей кэша")
        memory_size = sum(sys.getsizeof(v) for v in self._memory_cache.values())
        if memory_size > self.max_memory_mb * 1024 * 1024:
            sorted_keys = sorted(self._access_count.items(), key=lambda x: x[1])
            keys_to_remove = [k for k, _ in sorted_keys[:len(sorted_keys)//5]]
            for key in keys_to_remove:
                if key in self._memory_cache:
                    del self._memory_cache[key]
                if key in self._cache_timestamps:
                    del self._cache_timestamps[key]
                if key in self._access_count:
                    del self._access_count[key]
            logger.debug(f"🗑️ Очищено {len(keys_to_remove)} редко используемых записей")

    def clear_cache(self, cache_type: Optional[str] = None):
        if cache_type:
            cache_dirs = {
                'tariffs': self.tariffs_cache_dir,
                'api': self.api_cache_dir,
                'calc': self.calc_cache_dir,
                'user_categories': self.user_categories_cache_dir
            }
            if cache_type in cache_dirs:
                for cache_file in cache_dirs[cache_type].glob("*.cache"):
                    cache_file.unlink()
            prefix = f"{cache_type}:"
            keys_to_remove = [k for k in self._memory_cache if k.startswith(prefix)]
            for key in keys_to_remove:
                del self._memory_cache[key]
                if key in self._cache_timestamps:
                    del self._cache_timestamps[key]
                if key in self._access_count:
                    del self._access_count[key]
            logger.info(f"🗑️ Кэш очищен: {cache_type}")
        else:
            self._memory_cache.clear()
            self._cache_timestamps.clear()
            self._access_count.clear()
            for cache_dir in [self.tariffs_cache_dir, self.api_cache_dir,
                             self.calc_cache_dir, self.user_categories_cache_dir]:
                for cache_file in cache_dir.glob("*.cache"):
                    cache_file.unlink()
            logger.info("🗑️ Весь кэш очищен")

    def get_cache_stats(self) -> Dict[str, Any]:
        memory_size = sum(sys.getsizeof(v) for v in self._memory_cache.values())
        return {
            'memory_entries': len(self._memory_cache),
            'memory_size_mb': memory_size / (1024 * 1024),
            'tariffs_cache_files': len(list(self.tariffs_cache_dir.glob("*.cache"))),
            'api_cache_files': len(list(self.api_cache_dir.glob("*.cache"))),
            'calc_cache_files': len(list(self.calc_cache_dir.glob("*.cache"))),
            'user_categories_cache_files': len(list(self.user_categories_cache_dir.glob("*.cache"))),
            'cache_ttl_seconds': self.cache_ttl,
            'max_memory_mb': self.max_memory_mb
        }


# ============================================================================
# БЛОК 4: КОНФИГУРАЦИИ API ЯНДЕКС МАРКЕТ
# ============================================================================

class YandexMarketAPIEndpoint(Enum):
    TARIFFS = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/tariffs"
    COMMISSIONS = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/offer-mapping-entries"
    DELIVERY_ZONES = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/delivery-zones"
    ORDERS = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/orders"
    STOCKS = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/stocks"
    OFFERS = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/offers"
    CATEGORIES = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/categories"
    DEEPSEEK_CHAT = "https://api.deepseek.com/v1/chat/completions"

@dataclass
class YandexMarketTariffData:
    category: str
    commission_rate: float
    min_commission: float
    last_mile_base: float
    last_mile_per_kg: float
    last_mile_per_km: float
    acquiring_fee: float
    return_fee: float
    penalty_rate: float
    penalty_time_hours: int
    fbo_multiplier: float
    fbp_multiplier: float
    storage_base_rate: float
    min_logistics: float
    last_updated: str = ""
    source: str = "default"
    api_response_raw: str = ""
    confidence: float = 1.0

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'YandexMarketTariffData':
        return cls(**data)

    def get_commission_percent(self) -> float:
        return self.commission_rate * 100

    def get_total_delivery_cost(self, weight_kg: float, distance_km: float) -> float:
        return max(
            self.last_mile_base + (weight_kg * self.last_mile_per_kg) + (distance_km * self.last_mile_per_km),
            self.min_logistics
        )

    def get_storage_cost(self, days: int, units: int) -> float:
        return self.storage_base_rate * days * units


# ============================================================================
# БЛОК 5: API МЕНЕДЖЕР ДЛЯ ЯНДЕКС МАРКЕТ
# ============================================================================

class APIRateLimiter:
    """Класс для управления частотой API запросов."""
    def __init__(self):
        self.last_request_time: Dict[str, float] = {}
        self.min_interval: Dict[str, float] = {
            'yandex_market': 1.0,
            'deepseek': 0.5
        }
        self.request_count: Dict[str, int] = {}
        self.reset_time: Dict[str, float] = {}

    def wait_if_needed(self, service: str):
        current_time = time.time()
        if service in self.request_count:
            if self.request_count[service] >= 100:
                if current_time < self.reset_time.get(service, 0):
                    wait_time = self.reset_time[service] - current_time
                    if wait_time > 0:
                        logger.info(f"⏳ Достигнут лимит запросов к {service}, ожидание {wait_time:.1f}с")
                        time.sleep(wait_time)
                self.request_count[service] = 0
        if service in self.last_request_time:
            elapsed = current_time - self.last_request_time[service]
            min_wait = self.min_interval.get(service, 0.5)
            if elapsed < min_wait:
                time.sleep(min_wait - elapsed)
        self.last_request_time[service] = current_time
        self.request_count[service] = self.request_count.get(service, 0) + 1
        if service not in self.reset_time:
            self.reset_time[service] = current_time + 60

    def get_stats(self) -> Dict[str, Any]:
        return {
            'last_request_time': self.last_request_time,
            'request_count': self.request_count,
            'reset_time': self.reset_time
        }


class YandexMarketAPIManager:
    """Менеджер для работы с API Яндекс Маркет."""
    def __init__(self, cache_manager: Optional[CacheManager] = None,
                 secure_data: Optional[SecureDataManager] = None):
        self.cache_manager = cache_manager or CacheManager()
        self.secure_data = secure_data or SecureDataManager()
        self.rate_limiter = APIRateLimiter()
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': f'FBS-Unit-Economy-Pro/{APP_VERSION}',
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        })
        self.progress_tracker = ProgressTracker()
        self._api_keys_cache: Dict[str, str] = {}
        self._load_api_keys()
        self.audit_logger = AuditLogger()
        self._default_tariffs_cache: Dict[str, Dict] = {}
        self._session_start = datetime.now()
        logger.info(f"✅ YandexMarketAPIManager инициализирован")

    def _load_api_keys(self):
        try:
            if self.secure_data.is_available():
                self._api_keys_cache = self.secure_data.get_all_api_keys()
                if self._api_keys_cache:
                    logger.info(f"🔑 Загружены API ключи для: {', '.join(self._api_keys_cache.keys())}")
            if not self._api_keys_cache:
                key_file = CONFIG_DIR / "api_keys.json"
                if key_file.exists():
                    try:
                        with open(key_file, 'r', encoding='utf-8') as f:
                            self._api_keys_cache = json.load(f)
                        logger.info(f"🔑 Загружены API ключи из файла")
                    except Exception as e:
                        logger.warning(f"⚠️ Ошибка загрузки API ключей из файла: {e}")
        except Exception as e:
            logger.warning(f"⚠️ Не удалось загрузить API ключи: {e}")
            self._api_keys_cache = {}

    def save_api_key(self, service: str, api_key: str) -> bool:
        if not api_key or not api_key.strip():
            logger.warning("⚠️ API ключ пустой, сохранение отменено")
            return False
        api_key = api_key.strip()
        self._api_keys_cache[service] = api_key
        if self.secure_data.is_available():
            success = self.secure_data.store_api_key(service, api_key)
            if success:
                self.audit_logger.log('save_api_key', {'service': service, 'method': 'secure'})
                logger.info(f"✅ API ключ для {service} сохранен в защищенное хранилище")
                return True
        try:
            key_file = CONFIG_DIR / "api_keys.json"
            with open(key_file, 'w', encoding='utf-8') as f:
                json.dump(self._api_keys_cache, f, indent=2, ensure_ascii=False)
            try:
                os.chmod(key_file, 0o600)
            except OSError:
                pass
            self.audit_logger.log('save_api_key', {'service': service, 'method': 'json_fallback'})
            logger.info(f"✅ API ключ для {service} сохранен в файл")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения API ключа: {e}")
            return False

    def get_api_key(self, service: str) -> Optional[str]:
        return self._api_keys_cache.get(service)

    def has_api_key(self, service: str) -> bool:
        return bool(self._api_keys_cache.get(service))

    def get_cached_tariffs(self) -> Optional[Dict[str, Dict]]:
        cache_key = "yandex_market_tariffs"
        return self.cache_manager.get('tariffs', cache_key)

    def save_tariffs_to_cache(self, tariffs: Dict[str, Dict]):
        cache_key = "yandex_market_tariffs"
        self.cache_manager.set('tariffs', cache_key, {
            'tariffs': tariffs,
            'marketplace': 'Яндекс Маркет',
            'cached_at': datetime.now().isoformat(),
            'version': APP_VERSION
        })
        logger.info(f"💾 Тарифы Яндекс Маркет сохранены в кэш")

    def load_user_categories(self, csv_content: str, column_mapping: Dict[str, str]) -> Dict[str, Dict]:
        """
        Загрузить пользовательские категории с тарифами из CSV с учетом маппинга колонок.
        """
        tariffs = {}
        try:
            df = pd.read_csv(io.StringIO(csv_content))
            # Переименовываем колонки согласно маппингу
            rename_dict = {v: k for k, v in column_mapping.items() if v in df.columns}
            if rename_dict:
                df = df.rename(columns=rename_dict)
            required_cols = ['category', 'commission_rate', 'min_commission', 'last_mile_base']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise ValueError(f"Отсутствуют обязательные колонки: {', '.join(missing_cols)}")
            for _, row in df.iterrows():
                category = str(row.get('category', 'default'))
                tariffs[category] = {
                    'commission_rate': float(row.get('commission_rate', 0.15)),
                    'min_commission': float(row.get('min_commission', 30.0)),
                    'last_mile_base': float(row.get('last_mile_base', 50.0)),
                    'last_mile_per_kg': float(row.get('last_mile_per_kg', 15.0)),
                    'last_mile_per_km': float(row.get('last_mile_per_km', 3.5)),
                    'acquiring_fee': float(row.get('acquiring_fee', 0.015)),
                    'return_fee': float(row.get('return_fee', 0.02)),
                    'penalty_rate': float(row.get('penalty_rate', 0.05)),
                    'penalty_time_hours': float(row.get('penalty_time_hours', 24)),
                    'fbo_multiplier': float(row.get('fbo_multiplier', 0.75)),
                    'fbp_multiplier': float(row.get('fbp_multiplier', 0.60)),
                    'storage_base_rate': float(row.get('storage_base_rate', 0.30)),
                    'min_logistics': float(row.get('min_logistics', 25.0)),
                    'source': 'user_categories_csv',
                    'last_updated': datetime.now().isoformat(),
                    'confidence': 1.0
                }
            logger.info(f"✅ Загружено {len(tariffs)} категорий из пользовательского CSV")
            self.audit_logger.log('load_user_categories', {'count': len(tariffs)})
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки пользовательских категорий: {e}")
            tariffs = {}
        return tariffs

    @retry_on_failure(max_retries=2, delay=2.0)
    def fetch_yandex_market_tariffs(self) -> Dict[str, Dict]:
        tariffs = {}
        api_key = self.get_api_key('yandex_market')
        campaign_id = self.get_api_key('yandex_campaign_id')
        if not api_key or not campaign_id:
            logger.warning("⚠️ API ключи Яндекс Маркет не найдены")
            return {}
        self.rate_limiter.wait_if_needed('yandex_market')
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        try:
            tariffs_url = YandexMarketAPIEndpoint.TARIFFS.value.format(campaign_id=campaign_id)
            logger.info(f"📡 Запрос тарифов Яндекс Маркет (кампания: {campaign_id})...")
            response = self.session.get(tariffs_url, headers=headers, timeout=30)
            if response.status_code == 200:
                data = response.json()
                logger.info("✅ Получены тарифы Яндекс Маркет")
                for item in data.get('tariffs', []):
                    category = item.get('category', 'default')
                    category_name = item.get('categoryName', category)
                    tariffs[category_name] = {
                        'commission_rate': float(item.get('commission', 14.5)) / 100,
                        'min_commission': float(item.get('minCommission', 35)),
                        'last_mile_base': float(item.get('deliveryBase', 55)),
                        'last_mile_per_kg': float(item.get('deliveryPerKg', 16)),
                        'last_mile_per_km': float(item.get('deliveryPerKm', 3.8)),
                        'acquiring_fee': float(item.get('acquiringFee', 0.015)),
                        'return_fee': float(item.get('returnFee', 0.025)),
                        'penalty_rate': float(item.get('penaltyRate', 0.07)),
                        'penalty_time_hours': int(item.get('penaltyTimeHours', 24)),
                        'fbo_multiplier': float(item.get('fboMultiplier', 0.80)),
                        'fbp_multiplier': float(item.get('fbpMultiplier', 0.65)),
                        'storage_base_rate': float(item.get('storageRate', 0.35)),
                        'min_logistics': float(item.get('minDeliveryCost', 30)),
                        'source': 'yandex_api',
                        'last_updated': datetime.now().isoformat(),
                        'confidence': 1.0,
                        'api_response_raw': json.dumps(item, ensure_ascii=False)
                    }
                if tariffs:
                    logger.info(f"✅ Загружено {len(tariffs)} категорий тарифов Яндекс Маркет через API")
                    self.audit_logger.log('fetch_yandex_tariffs', {'count': len(tariffs), 'status': 'success'})
            else:
                logger.error(f"❌ Яндекс Маркет API вернул статус {response.status_code}")
                logger.error(f"Ответ: {response.text[:500]}")
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки тарифов Яндекс Маркет: {e}")
            logger.exception(e)
        return tariffs

    def fetch_tariffs_via_deepseek(self) -> Dict[str, Dict]:
        tariffs = {}
        api_key = self.get_api_key('deepseek')
        if not api_key:
            logger.warning("⚠️ DeepSeek API ключ не найден")
            return {}
        self.rate_limiter.wait_if_needed('deepseek')
        try:
            prompt = """
            Ты эксперт по тарифам Яндекс Маркет с актуальными данными на 2026 год.
            Предоставь актуальные тарифы для Яндекс Маркет в формате строгого JSON без markdown-разметки.
            ВАЖНО: Верни ТОЛЬКО валидный JSON объект, без каких-либо пояснений.
            Формат ответа:
            {
                "categories": {
                    "default": {
                        "commission_rate": 0.145,
                        "min_commission": 35,
                        "last_mile_base": 55,
                        "last_mile_per_kg": 16,
                        "last_mile_per_km": 3.8,
                        "acquiring_fee": 0.015,
                        "return_fee": 0.025,
                        "penalty_rate": 0.07,
                        "penalty_time_hours": 24,
                        "fbo_multiplier": 0.80,
                        "fbp_multiplier": 0.65,
                        "storage_base_rate": 0.35,
                        "min_logistics": 30
                    },
                    "electronics": {
                        "commission_rate": 0.10,
                        "min_commission": 30,
                        "last_mile_base": 50,
                        "last_mile_per_kg": 15,
                        "last_mile_per_km": 3.5,
                        "acquiring_fee": 0.015,
                        "return_fee": 0.02,
                        "penalty_rate": 0.05,
                        "penalty_time_hours": 24,
                        "fbo_multiplier": 0.75,
                        "fbp_multiplier": 0.60,
                        "storage_base_rate": 0.30,
                        "min_logistics": 25
                    },
                    "clothing": {
                        "commission_rate": 0.16,
                        "min_commission": 25,
                        "last_mile_base": 45,
                        "last_mile_per_kg": 14,
                        "last_mile_per_km": 3.2,
                        "acquiring_fee": 0.015,
                        "return_fee": 0.018,
                        "penalty_rate": 0.08,
                        "penalty_time_hours": 24,
                        "fbo_multiplier": 0.70,
                        "fbp_multiplier": 0.55,
                        "storage_base_rate": 0.25,
                        "min_logistics": 22
                    }
                },
                "source": "deepseek_ai",
                "confidence": 0.85,
                "data_collection_date": "2026-01"
            }
            Укажи реальные актуальные тарифы для Яндекс Маркет на 2026 год.
            Учти последние изменения в тарифной политике.
            """
            headers = {
                'Authorization': f'Bearer {api_key}',
                'Content-Type': 'application/json'
            }
            payload = {
                'model': 'deepseek-chat',
                'messages': [
                    {
                        'role': 'system',
                        'content': 'Ты эксперт по тарифам Яндекс Маркет. Отвечай только валидным JSON без пояснений.'
                    },
                    {
                        'role': 'user',
                        'content': prompt
                    }
                ],
                'temperature': 0.1,
                'max_tokens': 4000,
                'response_format': {'type': 'json_object'}
            }
            logger.info(f"🤖 Отправка запроса к DeepSeek AI для получения тарифов Яндекс Маркет...")
            response = self.session.post(
                YandexMarketAPIEndpoint.DEEPSEEK_CHAT.value,
                headers=headers,
                json=payload,
                timeout=60
            )
            if response.status_code == 200:
                response_data = response.json()
                content_text = response_data['choices'][0]['message']['content']
                try:
                    content = json.loads(content_text)
                except json.JSONDecodeError:
                    json_match = re.search(r'\{.*\}', content_text, re.DOTALL)
                    if json_match:
                        content = json.loads(json_match.group())
                    else:
                        raise ValueError("Не удалось извлечь JSON из ответа DeepSeek")
                categories = content.get('categories', {})
                confidence = content.get('confidence', 0.8)
                for category, tariff_data in categories.items():
                    tariff_data['source'] = 'deepseek_ai'
                    tariff_data['last_updated'] = datetime.now().isoformat()
                    tariff_data['confidence'] = confidence
                    tariffs[category] = tariff_data
                if tariffs:
                    logger.info(f"✅ DeepSeek предоставил тарифы для {len(tariffs)} категорий Яндекс Маркет")
                    self.audit_logger.log('fetch_deepseek_tariffs', {'count': len(tariffs)})
            else:
                logger.error(f"❌ DeepSeek API вернул статус {response.status_code}")
                logger.error(f"Ответ: {response.text[:500]}")
        except Exception as e:
            logger.error(f"❌ Ошибка запроса к DeepSeek: {e}")
            logger.exception(e)
        return tariffs

    def get_tariffs(self, force_refresh: bool = False,
                   use_ai_fallback: bool = True,
                   user_categories_csv: Optional[str] = None,
                   user_tariffs: Optional[Dict[str, Dict]] = None,
                   column_mapping: Dict[str, str] = None,
                   preferred_source: str = 'combined') -> Dict[str, Dict]:
        """
        Основной метод получения тарифов с каскадным фолбэком.
        preferred_source: 'api', 'ai', 'combined'
        """
        if not force_refresh:
            cached = self.get_cached_tariffs()
            if cached:
                tariffs = cached.get('tariffs', {})
                if tariffs:
                    cached_time = cached.get('cached_at', '')
                    logger.info(f"📦 Использованы кэшированные тарифы Яндекс Маркет от {cached_time}")
                    return tariffs
        logger.info(f"🔄 Загрузка тарифов Яндекс Маркет (источник: {preferred_source})...")
        tariffs = {}
        # Шаг 1: API Яндекс Маркет (если preferred_source 'api' или 'combined')
        if preferred_source in ('api', 'combined'):
            try:
                tariffs = self.fetch_yandex_market_tariffs()
                if tariffs and preferred_source == 'api':
                    # если запрошен только API и он успешен, возвращаем
                    self.save_tariffs_to_cache(tariffs)
                    return tariffs
            except Exception as e:
                logger.error(f"❌ Ошибка API Яндекс Маркет: {e}")
                tariffs = {}
        # Шаг 2: DeepSeek AI (если preferred_source 'ai' или 'combined' и API не дал результатов)
        if (preferred_source in ('ai', 'combined')) and not tariffs:
            logger.info(f"🤖 Использую DeepSeek AI (preferred_source={preferred_source})")
            try:
                ai_tariffs = self.fetch_tariffs_via_deepseek()
                if ai_tariffs:
                    tariffs = ai_tariffs
                    logger.info(f"✅ Тарифы Яндекс Маркет получены через DeepSeek AI")
                    if preferred_source == 'ai':
                        self.save_tariffs_to_cache(tariffs)
                        return tariffs
            except Exception as e:
                logger.error(f"❌ DeepSeek также недоступен: {e}")
        # Шаг 3: Пользовательские категории из CSV с маппингом (всегда как fallback)
        if not tariffs and user_categories_csv and column_mapping:
            logger.info(f"📄 Использую пользовательские категории из CSV с маппингом")
            csv_tariffs = self.load_user_categories(user_categories_csv, column_mapping)
            if csv_tariffs:
                tariffs = csv_tariffs
                logger.info(f"✅ Тарифы Яндекс Маркет загружены из пользовательского CSV")
        # Шаг 4: Пользовательские тарифы
        if not tariffs and user_tariffs:
            logger.info(f"👤 Использую пользовательские тарифы")
            tariffs = user_tariffs
            for category in tariffs:
                tariffs[category]['source'] = 'user_input'
                tariffs[category]['last_updated'] = datetime.now().isoformat()
                tariffs[category]['confidence'] = 1.0
        # Шаг 5: Базовые значения (только если ничего не загрузилось)
        if not tariffs:
            logger.warning(f"⚠️ ВСЕ ИСТОЧНИКИ НЕДОСТУПНЫ. Использую базовые значения.")
            tariffs = {
                'default': {
                    'commission_rate': 0.145,
                    'min_commission': 35,
                    'last_mile_base': 55,
                    'last_mile_per_kg': 16,
                    'last_mile_per_km': 3.8,
                    'acquiring_fee': 0.015,
                    'return_fee': 0.025,
                    'penalty_rate': 0.07,
                    'penalty_time_hours': 24,
                    'fbo_multiplier': 0.80,
                    'fbp_multiplier': 0.65,
                    'storage_base_rate': 0.35,
                    'min_logistics': 30,
                    'source': 'default_fallback',
                    'last_updated': datetime.now().isoformat(),
                    'confidence': 0.5
                },
                'electronics': {
                    'commission_rate': 0.10,
                    'min_commission': 30,
                    'last_mile_base': 50,
                    'last_mile_per_kg': 15,
                    'last_mile_per_km': 3.5,
                    'acquiring_fee': 0.015,
                    'return_fee': 0.02,
                    'penalty_rate': 0.05,
                    'penalty_time_hours': 24,
                    'fbo_multiplier': 0.75,
                    'fbp_multiplier': 0.60,
                    'storage_base_rate': 0.30,
                    'min_logistics': 25,
                    'source': 'default_fallback',
                    'last_updated': datetime.now().isoformat(),
                    'confidence': 0.5
                },
                'clothing': {
                    'commission_rate': 0.16,
                    'min_commission': 25,
                    'last_mile_base': 45,
                    'last_mile_per_kg': 14,
                    'last_mile_per_km': 3.2,
                    'acquiring_fee': 0.015,
                    'return_fee': 0.018,
                    'penalty_rate': 0.08,
                    'penalty_time_hours': 24,
                    'fbo_multiplier': 0.70,
                    'fbp_multiplier': 0.55,
                    'storage_base_rate': 0.25,
                    'min_logistics': 22,
                    'source': 'default_fallback',
                    'last_updated': datetime.now().isoformat(),
                    'confidence': 0.5
                }
            }
            self.audit_logger.log('tariffs_fallback', {})
        if tariffs:
            self.save_tariffs_to_cache(tariffs)
        return tariffs

    def get_all_tariffs_as_dataframe(self) -> pd.DataFrame:
        tariffs = self.get_tariffs()
        if not tariffs:
            return pd.DataFrame()
        rows = []
        for category, data in tariffs.items():
            rows.append({
                'Категория': category,
                'Комиссия, %': round(data.get('commission_rate', 0) * 100, 2),
                'Мин. комиссия, ₽': data.get('min_commission', 0),
                'База Last Mile, ₽': data.get('last_mile_base', 0),
                'Last Mile за кг, ₽': data.get('last_mile_per_kg', 0),
                'Last Mile за км, ₽': data.get('last_mile_per_km', 0),
                'Эквайринг, %': round(data.get('acquiring_fee', 0) * 100, 2),
                'Возвраты, %': round(data.get('return_fee', 0) * 100, 2),
                'Штраф за просрочку, %': round(data.get('penalty_rate', 0) * 100, 2),
                'Время на передачу, ч': data.get('penalty_time_hours', 0),
                'Множитель FBO': data.get('fbo_multiplier', 0),
                'Множитель FBP': data.get('fbp_multiplier', 0),
                'Хранение, ₽/день': data.get('storage_base_rate', 0),
                'Мин. логистика, ₽': data.get('min_logistics', 0),
                'Источник': data.get('source', 'unknown'),
                'Уверенность': f"{data.get('confidence', 1.0)*100:.0f}%",
                'Обновлено': data.get('last_updated', '')[:19] if data.get('last_updated') else ''
            })
        return pd.DataFrame(rows)

    def test_api_connection(self) -> Dict[str, Any]:
        result = {
            'marketplace': 'Яндекс Маркет',
            'timestamp': datetime.now().isoformat(),
            'status': 'unknown',
            'response_time_ms': 0,
            'error': None,
            'has_api_key': False
        }
        start_time = time.time()
        try:
            has_key = self.has_api_key('yandex_market') and self.has_api_key('yandex_campaign_id')
            result['has_api_key'] = has_key
            if not has_key:
                result['status'] = 'no_api_key'
                result['error'] = 'API ключи Яндекс Маркет не настроены'
            else:
                tariffs = self.fetch_yandex_market_tariffs()
                result['status'] = 'success' if tariffs else 'empty_response'
        except Exception as e:
            result['status'] = 'error'
            result['error'] = str(e)
        result['response_time_ms'] = round((time.time() - start_time) * 1000, 2)
        return result


# ============================================================================
# БЛОК 6: ДАТАКЛАССЫ ДЛЯ РАСЧЕТОВ
# ============================================================================

@dataclass
class ProductData:
    artikul: str = ""
    brand: str = ""
    category: str = "default"
    selling_price: float = 0.0
    cogs: float = 0.0
    weight_kg: float = 0.0
    length_cm: float = 0.0
    width_cm: float = 0.0
    height_cm: float = 0.0
    warehouse_distance_km: float = 0.0
    daily_sales: int = 5
    stock_depth_days: int = 30
    packaging_cost: float = 0.0
    marketing_budget_per_unit: float = 0.0
    operator_hourly_rate: float = 300.0
    pick_pack_time_min: float = 5.0
    pallet_capacity: int = 100
    transport_cost_per_km: float = 20.0
    safety_stock_days: int = 7
    supplier_lead_time_days: int = 3
    has_night_shift: bool = False
    seasonal_coefficients: Dict[int, float] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ProductData':
        known_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered_data = {k: v for k, v in data.items() if k in known_fields}
        return cls(**filtered_data)

    def validate(self) -> List[str]:
        errors = []
        if self.selling_price <= 0:
            errors.append("Цена продажи должна быть больше нуля")
        if self.cogs <= 0:
            errors.append("Себестоимость должна быть больше нуля")
        if self.cogs >= self.selling_price:
            errors.append("Себестоимость не может быть больше или равна цене продажи")
        if self.weight_kg < 0:
            errors.append("Вес не может быть отрицательным")
        if self.daily_sales <= 0:
            errors.append("Продажи в день должны быть больше нуля")
        if self.stock_depth_days <= 0:
            errors.append("Глубина запаса должна быть больше нуля")
        return errors

    def get_volume_weight(self) -> float:
        if self.length_cm > 0 and self.width_cm > 0 and self.height_cm > 0:
            return (self.length_cm * self.width_cm * self.height_cm) / 5000.0
        return 0.0

    def get_billable_weight(self) -> float:
        return max(self.weight_kg, self.get_volume_weight())


@dataclass
class CalculationResult:
    artikul: str = ""
    brand: str = ""
    category: str = ""
    selling_price: float = 0.0
    total_expenses: float = 0.0
    gross_profit: float = 0.0
    margin_percent: float = 0.0
    roi_percent: float = 0.0
    commission: float = 0.0
    first_mile_cost: float = 0.0
    last_mile_cost: float = 0.0
    pick_pack_cost: float = 0.0
    packaging_cost: float = 0.0
    acquiring_cost: float = 0.0
    return_cost: float = 0.0
    penalty_cost: float = 0.0
    marketing_cost: float = 0.0
    warehouse_cost: float = 0.0
    tax_cost: float = 0.0
    penalty_probability: float = 0.0
    break_even_distance_km: float = 0.0
    max_discount_percent: float = 0.0
    safety_margin_price: float = 0.0
    break_even_volume: float = 0.0
    optimal_stock_units: int = 0
    safety_stock_units: int = 0
    reorder_point_units: int = 0
    stock_turnover_days: float = 0.0
    stock_turnover_rate: float = 0.0
    logistic_zone: str = "unknown"
    logistic_zone_label: str = ""
    logistic_recommendation: str = ""
    is_logistic_critical: bool = False
    space_efficiency_ratio: float = 0.0
    revenue_per_sqm: float = 0.0
    profit_per_sqm: float = 0.0
    seasonal_factor: float = 1.0
    adjusted_margin_percent: float = 0.0
    data_source: str = "unknown"
    data_confidence: float = 1.0
    optimal_price: float = 0.0

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    def get_summary(self) -> Dict[str, Any]:
        return {
            'Артикул': self.artikul,
            'Бренд': self.brand,
            'Категория': self.category,
            'Цена, ₽': self.selling_price,
            'Прибыль, ₽': self.gross_profit,
            'Маржа, %': self.margin_percent,
            'ROI, %': self.roi_percent,
            'First Mile, ₽': self.first_mile_cost,
            'Last Mile, ₽': self.last_mile_cost,
            'Опт. запас, шт': self.optimal_stock_units,
            'Оборачиваемость, дн': self.stock_turnover_days,
            'Лог. зона': self.logistic_zone_label,
            'Скорр. маржа, %': self.adjusted_margin_percent,
            'Источник': self.data_source
        }

    def is_profitable(self) -> bool:
        return self.gross_profit > 0


# ============================================================================
# БЛОК 7: КОНФИГУРАЦИИ НАЛОГОВЫХ СИСТЕМ
# ============================================================================

TAX_SYSTEMS = {
    "УСН 6% (доходы)": {
        "rate": 0.06,
        "base": "revenue",
        "name": "УСН_6",
        "description": "Упрощенная система налогообложения, 6% от доходов"
    },
    "УСН 15% (доходы-расходы)": {
        "rate": 0.15,
        "base": "profit",
        "min_rate": 0.01,
        "name": "УСН_15",
        "description": "Упрощенная система налогообложения, 15% от прибыли (мин. 1% от доходов)"
    },
    "ОСН (общая)": {
        "rate": 0.20,
        "base": "profit",
        "name": "ОСН",
        "description": "Общая система налогообложения, 20% налог на прибыль"
    },
    "НПД (самозанятый)": {
        "rate": 0.06,
        "base": "revenue",
        "name": "НПД",
        "description": "Налог на профессиональный доход, 6% от доходов с юрлицами"
    },
    "Патент": {
        "rate": 0.06,
        "base": "revenue",
        "name": "Патент",
        "description": "Патентная система налогообложения"
    }
}


# ============================================================================
# БЛОК 8: КАЛЬКУЛЯТОР FBS ДЛЯ ЯНДЕКС МАРКЕТ
# ============================================================================

class YandexMarketCalculator:
    def __init__(self, api_manager: Optional[YandexMarketAPIManager] = None,
                 tax_system: str = "УСН 6% (доходы)"):
        self.api_manager = api_manager or YandexMarketAPIManager()
        self.tax_system = tax_system
        self.current_tariffs: Dict[str, Dict] = {}
        self.tariffs_updated_at: Optional[datetime] = None
        self.tariffs_source = "unknown"
        self.user_categories: Dict[str, Dict] = {}
        self.progress_tracker = ProgressTracker()
        self.audit_logger = AuditLogger()
        self._load_tariffs()
        logger.info("✅ YandexMarketCalculator инициализирован")

    def _load_tariffs(self):
        self.current_tariffs = self.api_manager.get_tariffs()
        self.tariffs_updated_at = datetime.now()
        sources = set()
        for tariff in self.current_tariffs.values():
            source = tariff.get('source', 'unknown')
            if 'yandex_api' in source:
                sources.add('api')
            elif 'deepseek' in source:
                sources.add('deepseek')
            elif 'user_categories' in source:
                sources.add('user_categories')
            elif 'user_input' in source:
                sources.add('user')
            else:
                sources.add('unknown')
        if 'api' in sources:
            self.tariffs_source = 'api'
        elif 'deepseek' in sources:
            self.tariffs_source = 'deepseek'
        elif 'user_categories' in sources:
            self.tariffs_source = 'user_categories'
        elif 'user' in sources:
            self.tariffs_source = 'user'
        else:
            self.tariffs_source = 'unknown'
        logger.info(f"✅ Тарифы загружены. Источник: {self.tariffs_source}. Категорий: {len(self.current_tariffs)}")

    def refresh_tariffs(self, force: bool = False, use_ai: bool = False,
                       user_categories_csv: Optional[str] = None,
                       user_tariffs: Optional[Dict[str, Dict]] = None,
                       column_mapping: Dict[str, str] = None,
                       preferred_source: str = 'combined'):
        """
        Обновить тарифы с выбором источника.
        preferred_source: 'api', 'ai', 'combined'
        """
        logger.info(f"🔄 Обновление тарифов Яндекс Маркет (источник: {preferred_source})...")
        # Для обратной совместимости: если use_ai=True и preferred_source не указан, устанавливаем 'combined'
        if use_ai and preferred_source == 'combined':
            pass  # уже combined
        elif use_ai and preferred_source == 'api':
            preferred_source = 'combined'  # если просят API с AI, то по сути combined
        self.current_tariffs = self.api_manager.get_tariffs(
            force_refresh=force,
            use_ai_fallback=use_ai,
            user_categories_csv=user_categories_csv,
            user_tariffs=user_tariffs,
            column_mapping=column_mapping,
            preferred_source=preferred_source
        )
        self.tariffs_updated_at = datetime.now()
        sources = set()
        for tariff in self.current_tariffs.values():
            source = tariff.get('source', 'unknown')
            if 'yandex_api' in source:
                sources.add('api')
            elif 'deepseek' in source:
                sources.add('deepseek')
            elif 'user_categories' in source:
                sources.add('user_categories')
            elif 'user_input' in source:
                sources.add('user')
            else:
                sources.add('unknown')
        if 'api' in sources:
            self.tariffs_source = 'api'
        elif 'deepseek' in sources:
            self.tariffs_source = 'deepseek'
        elif 'user_categories' in sources:
            self.tariffs_source = 'user_categories'
        elif 'user' in sources:
            self.tariffs_source = 'user'
        else:
            self.tariffs_source = 'unknown'
        logger.info(f"✅ Тарифы обновлены. Источник: {self.tariffs_source}. Категорий: {len(self.current_tariffs)}")

    def get_tariff_for_category(self, category: str) -> Dict[str, Any]:
        if not self.current_tariffs:
            logger.warning("⚠️ Тарифы не загружены. Загрузите тарифы через API, CSV или введите вручную.")
            return {}
        if category in self.current_tariffs:
            return self.current_tariffs[category]
        category_lower = category.lower()
        for cat, tariff in self.current_tariffs.items():
            cat_lower = cat.lower()
            if category_lower in cat_lower or cat_lower in category_lower:
                logger.debug(f"🔍 Найдено частичное совпадение: {category} -> {cat}")
                return tariff
        if 'default' in self.current_tariffs:
            logger.debug(f"🔍 Использую категорию 'default' для: {category}")
            return self.current_tariffs['default']
        first_tariff = next(iter(self.current_tariffs.values()), {})
        if first_tariff:
            logger.warning(f"⚠️ Категория {category} не найдена в тарифах, использую первую доступную")
            return first_tariff
        return {}

    def _get_logistic_zone(self, distance_km: float) -> Dict[str, Any]:
        if distance_km <= 25:
            return {
                'zone': 'red',
                'label': '🔴 Критическая зона',
                'recommendation': 'Срочно оптимизировать логистику! Рассмотрите смену поставщика или увеличение загрузки паллет.',
                'is_critical': True
            }
        elif distance_km <= 50:
            return {
                'zone': 'yellow',
                'label': '🟡 Зона риска',
                'recommendation': 'Рассмотрите смену поставщика логистики или оптимизацию маршрутов.',
                'is_critical': False
            }
        elif distance_km <= 100:
            return {
                'zone': 'green',
                'label': '🟢 Безопасная зона',
                'recommendation': 'Логистика оптимальна. Продолжайте мониторинг.',
                'is_critical': False
            }
        else:
            return {
                'zone': 'blue',
                'label': '🔵 Идеальная зона',
                'recommendation': 'Отличная логистическая стратегия! Рекомендуется масштабирование.',
                'is_critical': False
            }

    @timing_decorator
    def calculate_unit(self, product: ProductData) -> CalculationResult:
        validation_errors = product.validate()
        if validation_errors:
            logger.warning(f"⚠️ Ошибки валидации для {product.artikul}: {validation_errors}")
        result = CalculationResult()
        result.artikul = product.artikul
        result.brand = product.brand
        result.category = product.category
        result.selling_price = product.selling_price
        tariff = self.get_tariff_for_category(product.category)
        if not tariff:
            logger.error(f"❌ Тариф не найден для категории {product.category}.")
            tariff = {
                'commission_rate': 0.145, 'min_commission': 35,
                'last_mile_base': 55, 'last_mile_per_kg': 16,
                'acquiring_fee': 0.015, 'return_fee': 0.025,
                'penalty_rate': 0.07, 'penalty_time_hours': 24,
                'fbo_multiplier': 0.80, 'fbp_multiplier': 0.65,
                'storage_base_rate': 0.35, 'min_logistics': 30,
                'source': 'not_loaded', 'confidence': 0.0
            }
        result.data_source = tariff.get('source', 'unknown')
        result.data_confidence = tariff.get('confidence', 1.0)
        # 1. Комиссия Яндекс Маркет
        commission_rate = tariff.get('commission_rate', 0.145)
        min_commission = tariff.get('min_commission', 35)
        result.commission = max(product.selling_price * commission_rate, min_commission)
        # 2. First Mile
        if product.warehouse_distance_km > 0 and product.pallet_capacity > 0:
            cost_per_pallet = product.warehouse_distance_km * product.transport_cost_per_km * 2
            result.first_mile_cost = cost_per_pallet / product.pallet_capacity
        else:
            result.first_mile_cost = 0
        # 3. Last Mile
        billable_weight = product.get_billable_weight()
        billable_weight = math.ceil(billable_weight * 2) / 2
        last_mile_base = tariff.get('last_mile_base', 55)
        last_mile_per_kg = tariff.get('last_mile_per_kg', 16)
        min_logistics = tariff.get('min_logistics', 30)
        result.last_mile_cost = max(last_mile_base + (billable_weight * last_mile_per_kg), min_logistics)
        # 4. Pick & Pack
        pick_pack_hours = product.pick_pack_time_min / 60.0
        result.pick_pack_cost = pick_pack_hours * product.operator_hourly_rate
        # 5. Упаковка
        result.packaging_cost = product.packaging_cost
        # 6. Эквайринг
        acquiring_fee = tariff.get('acquiring_fee', 0.015)
        result.acquiring_cost = product.selling_price * acquiring_fee
        # 7. Возвраты
        return_fee = tariff.get('return_fee', 0.025)
        result.return_cost = product.selling_price * return_fee
        # 8. Штрафы за просрочку
        if product.has_night_shift:
            penalty_probability = 0.05
        else:
            penalty_probability = 0.35
        penalty_rate = tariff.get('penalty_rate', 0.07)
        result.penalty_probability = penalty_probability
        result.penalty_cost = product.selling_price * penalty_rate * penalty_probability
        # 9. Маркетинг
        result.marketing_cost = product.marketing_budget_per_unit
        # 10. Складские расходы
        total_stock = product.stock_depth_days * product.daily_sales
        if total_stock > 0 and product.daily_sales > 0:
            warehouse_space = 0.01  # м² на единицу
            total_warehouse_space = warehouse_space * total_stock
            warehouse_rent_per_sqm = 500  # ₽/м²
            monthly_rent = warehouse_rent_per_sqm * total_warehouse_space
            result.warehouse_cost = monthly_rent / (30 * product.daily_sales)
        else:
            result.warehouse_cost = 0
        # 11. Налог
        tax_config = TAX_SYSTEMS.get(self.tax_system, TAX_SYSTEMS["УСН 6% (доходы)"])
        if tax_config["base"] == "revenue":
            result.tax_cost = product.selling_price * tax_config["rate"]
        else:
            pre_tax_expenses = (
                result.commission + result.first_mile_cost + result.last_mile_cost +
                result.pick_pack_cost + result.packaging_cost + result.acquiring_cost +
                result.return_cost + result.penalty_cost + result.marketing_cost +
                result.warehouse_cost + product.cogs
            )
            pre_tax_profit = product.selling_price - pre_tax_expenses
            result.tax_cost = max(0, pre_tax_profit * tax_config["rate"])
            if "min_rate" in tax_config:
                min_tax = product.selling_price * tax_config["min_rate"]
                result.tax_cost = max(result.tax_cost, min_tax)
        # 12. Итого расходов и прибыль
        result.total_expenses = (
            product.cogs + result.commission + result.first_mile_cost +
            result.last_mile_cost + result.pick_pack_cost + result.packaging_cost +
            result.acquiring_cost + result.return_cost + result.penalty_cost +
            result.marketing_cost + result.warehouse_cost + result.tax_cost
        )
        result.gross_profit = result.selling_price - result.total_expenses
        if result.selling_price > 0:
            result.margin_percent = (result.gross_profit / result.selling_price) * 100
        else:
            result.margin_percent = 0
        if product.cogs > 0:
            result.roi_percent = (result.gross_profit / product.cogs) * 100
        else:
            result.roi_percent = 0
        # 13. Точка безубыточности по расстоянию
        if result.first_mile_cost > 0 and product.pallet_capacity > 0:
            cost_per_km_per_unit = (product.transport_cost_per_km * 2) / product.pallet_capacity
            if cost_per_km_per_unit > 0:
                result.break_even_distance_km = result.gross_profit / cost_per_km_per_unit
            else:
                result.break_even_distance_km = float('inf')
        else:
            result.break_even_distance_km = float('inf')
        # 14. Логистические зоны риска
        logistic_zone_info = self._get_logistic_zone(
            result.break_even_distance_km if result.break_even_distance_km != float('inf') else 200
        )
        result.logistic_zone = logistic_zone_info['zone']
        result.logistic_zone_label = logistic_zone_info['label']
        result.logistic_recommendation = logistic_zone_info['recommendation']
        result.is_logistic_critical = logistic_zone_info['is_critical']
        # 15. Запас прочности по цене
        variable_costs_percent = (
            commission_rate + acquiring_fee + return_fee +
            penalty_rate * penalty_probability +
            TAX_SYSTEMS[self.tax_system]["rate"]
        )
        fixed_costs_per_unit = (
            product.cogs + result.first_mile_cost + result.last_mile_cost +
            result.pick_pack_cost + result.packaging_cost +
            result.marketing_cost + result.warehouse_cost
        )
        denominator = 1 - variable_costs_percent
        if denominator > 0:
            min_price = fixed_costs_per_unit / denominator
        else:
            min_price = fixed_costs_per_unit * 2
        result.safety_margin_price = product.selling_price - min_price
        if product.selling_price > 0:
            result.max_discount_percent = ((product.selling_price - min_price) / product.selling_price) * 100
        else:
            result.max_discount_percent = 0
        # 16. Точка безубыточности по объему
        variable_costs = result.commission + result.last_mile_cost + result.acquiring_cost + result.return_cost + result.penalty_cost
        if (result.selling_price - variable_costs) > 0:
            result.break_even_volume = fixed_costs_per_unit / (result.selling_price - variable_costs)
        else:
            result.break_even_volume = float('inf')
        # 17. Сезонная корректировка
        current_month = datetime.now().month
        if product.seasonal_coefficients:
            seasonal_factor = product.seasonal_coefficients.get(current_month, 1.0)
        else:
            seasonal_factor = 1.0
        result.seasonal_factor = seasonal_factor
        result.adjusted_margin_percent = result.margin_percent * seasonal_factor
        # 18. Оптимизация складских остатков (EOQ)
        daily_demand = product.daily_sales
        annual_demand = daily_demand * 365
        ordering_cost = 500.0
        holding_cost_per_unit = 0.30
        if holding_cost_per_unit > 0 and ordering_cost > 0:
            eoq = math.sqrt((2 * annual_demand * ordering_cost) / holding_cost_per_unit)
            result.optimal_stock_units = int(math.ceil(eoq))
        else:
            result.optimal_stock_units = product.stock_depth_days * daily_demand
        max_daily_demand = daily_demand * 1.5
        result.safety_stock_units = int(math.ceil(
            (max_daily_demand * product.supplier_lead_time_days) -
            (daily_demand * product.supplier_lead_time_days)
        ))
        result.reorder_point_units = int(math.ceil(
            (daily_demand * product.supplier_lead_time_days) + result.safety_stock_units
        ))
        if daily_demand > 0 and result.optimal_stock_units > 0:
            result.stock_turnover_days = result.optimal_stock_units / daily_demand
        else:
            result.stock_turnover_days = 0
        if result.optimal_stock_units > 0:
            result.stock_turnover_rate = annual_demand / result.optimal_stock_units
        else:
            result.stock_turnover_rate = 0
        # 19. Эффективность использования пространства
        total_stock = product.stock_depth_days * product.daily_sales
        if total_stock > 0:
            total_sqm = total_stock * 0.01
            if total_sqm > 0:
                result.space_efficiency_ratio = total_stock / total_sqm
                result.revenue_per_sqm = (product.selling_price * product.daily_sales * 30) / total_sqm
                result.profit_per_sqm = (result.gross_profit * product.daily_sales * 30) / total_sqm
        # 20. Оптимальная цена
        result.optimal_price = product.selling_price
        self.audit_logger.log('calculate_unit', {
            'artikul': product.artikul,
            'profit': result.gross_profit,
            'margin': result.margin_percent,
            'logistic_zone': result.logistic_zone,
            'data_source': result.data_source
        })
        return result

    @timing_decorator
    def calculate_batch(self, products: List[ProductData],
                       use_parallel: bool = True,
                       max_workers: int = 8) -> List[CalculationResult]:
        total = len(products)
        results = [None] * total
        if total == 0:
            return []
        self.progress_tracker.start(total, f"Расчет {total} товаров...")
        if total > 100 and use_parallel:
            logger.info(f"⚡ Запуск параллельной обработки {total} товаров ({max_workers} потоков)")
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_index = {}
                for i, product in enumerate(products):
                    future = executor.submit(self.calculate_unit, product)
                    future_to_index[future] = i
                completed = 0
                for future in as_completed(future_to_index):
                    index = future_to_index[future]
                    try:
                        result = future.result()
                        results[index] = result
                        completed += 1
                        if completed % 100 == 0 or completed == total:
                            self.progress_tracker.update(completed, f"Обработано {completed}/{total} товаров")
                    except Exception as e:
                        logger.error(f"❌ Ошибка расчета товара с индексом {index}: {e}")
                        error_result = CalculationResult()
                        error_result.artikul = f"ERROR_{index}"
                        error_result.brand = f"Ошибка: {str(e)[:50]}"
                        results[index] = error_result
                        completed += 1
        else:
            logger.info(f"🔄 Запуск последовательной обработки {total} товаров")
            for i, product in enumerate(products):
                try:
                    results[i] = self.calculate_unit(product)
                except Exception as e:
                    logger.error(f"❌ Ошибка расчета товара {product.artikul}: {e}")
                    error_result = CalculationResult()
                    error_result.artikul = product.artikul
                    error_result.brand = f"Ошибка: {str(e)[:50]}"
                    results[i] = error_result
                if (i + 1) % 50 == 0 or (i + 1) == total:
                    self.progress_tracker.update(i + 1, f"Обработано {i + 1}/{total} товаров")
        self.progress_tracker.update(total, f"✅ Расчет завершен! Обработано {total} товаров")
        results = [r for r in results if r is not None]
        logger.info(f"✅ Пакетный расчет завершен. Успешно: {len(results)}/{total}")
        return results

    def generate_recommendations(self, results: List[CalculationResult]) -> List[Dict]:
        recommendations = []
        if not results:
            return recommendations
        total = len(results)
        profitable = [r for r in results if r.gross_profit > 0]
        unprofitable = [r for r in results if r.gross_profit <= 0]
        if len(unprofitable) / total > 0.2:
            recommendations.append({
                'priority': 'high',
                'category': 'Прибыльность',
                'icon': '⚠️',
                'message': f'{len(unprofitable)} товаров ({len(unprofitable)/total*100:.0f}%) убыточны. Рекомендуется пересмотреть цены или сократить затраты First Mile.',
                'affected_products': [r.artikul for r in unprofitable[:5]]
            })
        elif len(unprofitable) > 0:
            recommendations.append({
                'priority': 'medium',
                'category': 'Прибыльность',
                'icon': '📊',
                'message': f'{len(unprofitable)} товаров убыточны. Рассмотрите возможность повышения цен или оптимизации затрат.',
                'affected_products': [r.artikul for r in unprofitable[:5]]
            })
        high_distance = [r for r in results if r.is_logistic_critical and r.break_even_distance_km < 25]
        if high_distance:
            recommendations.append({
                'priority': 'high',
                'category': 'Логистика',
                'icon': '🚚',
                'message': f'{len(high_distance)} товаров находятся в критической логистической зоне (<25 км). Рекомендуется оптимизировать маршруты или увеличить загрузку паллет.',
                'affected_products': [r.artikul for r in high_distance[:5]]
            })
        high_stock = [r for r in results if r.stock_turnover_days > 60]
        if high_stock:
            recommendations.append({
                'priority': 'medium',
                'category': 'Склад',
                'icon': '📦',
                'message': f'{len(high_stock)} товаров имеют низкую оборачиваемость (>60 дней). Рекомендуется провести акции для ускорения оборота.',
                'affected_products': [r.artikul for r in high_stock[:5]]
            })
        low_confidence = [r for r in results if r.data_confidence < 0.9]
        if low_confidence:
            recommendations.append({
                'priority': 'info',
                'category': 'Качество данных',
                'icon': '🔍',
                'message': f'{len(low_confidence)} товаров используют данные с пониженной уверенностью. Рекомендуется загрузить точные тарифы через API или CSV.',
                'affected_products': [r.artikul for r in low_confidence[:5]]
            })
        self.audit_logger.log('generate_recommendations', {'count': len(recommendations)})
        return recommendations


# ============================================================================
# БЛОК 9: ЭКСПОРТ В EXCEL С ЖИВЫМИ ФОРМУЛАМИ И ТРЕМЯ ЛИСТАМИ
# ============================================================================

class ExcelExporter:
    @staticmethod
    def export_with_formulas(
        results: List[CalculationResult],
        products: List[ProductData],
        tariffs: Dict[str, Dict],
        tax_system: str
    ) -> bytes:
        if not OPENPYXL_AVAILABLE:
            raise ImportError("OpenPyXL не установлен")
        wb = Workbook()
        # -------------------- ЛИСТ 1: ТАРИФЫ --------------------
        ws_tariffs = wb.active
        ws_tariffs.title = "Тарифы"
        tariff_headers = ['Категория', 'Комиссия,%', 'Мин.комиссия', 'Last Mile база',
                         'Last Mile за кг', 'Эквайринг,%', 'Возвраты,%', 'Штрафы,%',
                         'Источник', 'Обновлено']
        for col, header in enumerate(tariff_headers, 1):
            cell = ws_tariffs.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1a1a2e", fill_type="solid")
        row = 2
        for category, tariff in tariffs.items():
            ws_tariffs.cell(row=row, column=1, value=category)
            ws_tariffs.cell(row=row, column=2, value=round(tariff.get('commission_rate', 0) * 100, 2))
            ws_tariffs.cell(row=row, column=3, value=tariff.get('min_commission', 0))
            ws_tariffs.cell(row=row, column=4, value=tariff.get('last_mile_base', 0))
            ws_tariffs.cell(row=row, column=5, value=tariff.get('last_mile_per_kg', 0))
            ws_tariffs.cell(row=row, column=6, value=round(tariff.get('acquiring_fee', 0) * 100, 2))
            ws_tariffs.cell(row=row, column=7, value=round(tariff.get('return_fee', 0) * 100, 2))
            ws_tariffs.cell(row=row, column=8, value=round(tariff.get('penalty_rate', 0) * 100, 2))
            ws_tariffs.cell(row=row, column=9, value=tariff.get('source', 'unknown'))
            ws_tariffs.cell(row=row, column=10, value=tariff.get('last_updated', '')[:19])
            row += 1
        # -------------------- ЛИСТ 2: РАСЧЁТЫ --------------------
        ws_results = wb.create_sheet("Расчёты")
        headers = ['Артикул', 'Бренд', 'Категория', 'Цена, ₽', 'Себестоимость, ₽',
                  'Комиссия, ₽', 'First Mile, ₽', 'Last Mile, ₽', 'Pick&Pack, ₽',
                  'Упаковка, ₽', 'Эквайринг, ₽', 'Возвраты, ₽', 'Штрафы, ₽',
                  'Маркетинг, ₽', 'Склад, ₽', 'Налог, ₽',
                  'Итого расходов, ₽', 'Прибыль, ₽', 'Маржа, %', 'ROI, %',
                  'Опт. запас, шт', 'Оборачиваемость, дн', 'Лог. зона', 'Источник']
        for col, header in enumerate(headers, 1):
            cell = ws_results.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1a1a2e", fill_type="solid")
        for i, (result, product) in enumerate(zip(results, products)):
            row = i + 2
            ws_results.cell(row=row, column=1, value=result.artikul)
            ws_results.cell(row=row, column=2, value=result.brand)
            ws_results.cell(row=row, column=3, value=result.category)
            ws_results.cell(row=row, column=4, value=product.selling_price)
            ws_results.cell(row=row, column=5, value=product.cogs)
            # Комиссия (формула)
            commission_rate = tariffs.get(result.category, {}).get('commission_rate', 0.145)
            ws_results.cell(row=row, column=6, value=result.commission)
            ws_results.cell(row=row, column=7, value=result.first_mile_cost)
            ws_results.cell(row=row, column=8, value=result.last_mile_cost)
            ws_results.cell(row=row, column=9, value=result.pick_pack_cost)
            ws_results.cell(row=row, column=10, value=result.packaging_cost)
            ws_results.cell(row=row, column=11, value=result.acquiring_cost)
            ws_results.cell(row=row, column=12, value=result.return_cost)
            ws_results.cell(row=row, column=13, value=result.penalty_cost)
            ws_results.cell(row=row, column=14, value=result.marketing_cost)
            ws_results.cell(row=row, column=15, value=result.warehouse_cost)
            ws_results.cell(row=row, column=16, value=result.tax_cost)
            # Итого расходов (формула)
            start_col = 5
            end_col = 16
            formula = f"=SUM({get_column_letter(start_col)}{row}:{get_column_letter(end_col)}{row})"
            ws_results.cell(row=row, column=17, value=formula)
            # Прибыль (формула)
            ws_results.cell(row=row, column=18, value=f"=D{row}-Q{row}")
            # Маржа (формула)
            ws_results.cell(row=row, column=19, value=f"=R{row}/D{row}*100")
            # ROI (формула)
            ws_results.cell(row=row, column=20, value=f"=R{row}/E{row}*100")
            ws_results.cell(row=row, column=21, value=result.optimal_stock_units)
            ws_results.cell(row=row, column=22, value=result.stock_turnover_days)
            ws_results.cell(row=row, column=23, value=result.logistic_zone_label)
            ws_results.cell(row=row, column=24, value=result.data_source)
            # Условное форматирование прибыли
            if result.gross_profit > 0:
                ws_results.cell(row=row, column=18).fill = PatternFill(start_color="C6EFCE", fill_type="solid")
            else:
                ws_results.cell(row=row, column=18).fill = PatternFill(start_color="FFC7CE", fill_type="solid")
        # Итоговая строка
        last_row = len(results) + 2
        if len(results) > 0:
            total_row = last_row + 1
            ws_results.cell(row=total_row, column=1, value="ИТОГО:")
            ws_results.cell(row=total_row, column=1).font = Font(bold=True)
            for col in [4, 5, 17, 18, 19, 20]:
                ws_results.cell(row=total_row, column=col, value=f"=SUM({get_column_letter(col)}2:{get_column_letter(col)}{last_row})")
                ws_results.cell(row=total_row, column=col).font = Font(bold=True)
        # Цветовая шкала для маржи
        if len(results) > 0:
            ws_results.conditional_formatting.add(
                f"S2:S{last_row}",
                ColorScaleRule(start_type="min", start_color="FFC7CE",
                             mid_type="percentile", mid_value=50, mid_color="FFEB9C",
                             end_type="max", end_color="C6EFCE")
            )
        # Ширина колонок
        for col in range(1, 25):
            ws_results.column_dimensions[get_column_letter(col)].width = 15

        # -------------------- ЛИСТ 3: ABC/XYZ + ДАШБОРДЫ --------------------
        ws_abc = wb.create_sheet("ABC_XYZ_Dashboard")
        # Заголовки
        abc_headers = ['Артикул', 'Прибыль, ₽', 'Доля прибыли, %', 'ABC класс',
                       'Выручка, ₽', 'Доля выручки, %', 'ABC класс (выручка)',
                       'Daily Sales', 'Среднее', 'Стд. отклонение', 'Коэф. вариации', 'XYZ класс']
        for col, header in enumerate(abc_headers, 1):
            cell = ws_abc.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1a1a2e", fill_type="solid")
        # Данные для ABC/XYZ
        total_profit = sum(r.gross_profit for r in results) if results else 1
        total_revenue = sum(r.selling_price for r in results) if results else 1
        for i, result in enumerate(results):
            row = i + 2
            ws_abc.cell(row=row, column=1, value=result.artikul)
            ws_abc.cell(row=row, column=2, value=result.gross_profit)
            profit_share = (result.gross_profit / total_profit * 100) if total_profit != 0 else 0
            ws_abc.cell(row=row, column=3, value=profit_share)
            # ABC по прибыли
            if profit_share >= 60:
                ws_abc.cell(row=row, column=4, value="A")
            elif profit_share >= 30:
                ws_abc.cell(row=row, column=4, value="B")
            else:
                ws_abc.cell(row=row, column=4, value="C")
            ws_abc.cell(row=row, column=5, value=result.selling_price)
            revenue_share = (result.selling_price / total_revenue * 100) if total_revenue != 0 else 0
            ws_abc.cell(row=row, column=6, value=revenue_share)
            # ABC по выручке
            if revenue_share >= 60:
                ws_abc.cell(row=row, column=7, value="A")
            elif revenue_share >= 30:
                ws_abc.cell(row=row, column=7, value="B")
            else:
                ws_abc.cell(row=row, column=7, value="C")
            # XYZ (симулируем на основе daily_sales)
            daily_sales = 5  # по умолчанию
            if i < len(products):
                daily_sales = products[i].daily_sales
            ws_abc.cell(row=row, column=8, value=daily_sales)
            # Имитация среднего и std для демонстрации (в реальности нужны исторические данные)
            avg = daily_sales * 1.0
            std = daily_sales * 0.3  # допущение
            ws_abc.cell(row=row, column=9, value=avg)
            ws_abc.cell(row=row, column=10, value=std)
            cv = (std / avg * 100) if avg != 0 else 0
            ws_abc.cell(row=row, column=11, value=cv)
            if cv < 20:
                ws_abc.cell(row=row, column=12, value="X")
            elif cv < 50:
                ws_abc.cell(row=row, column=12, value="Y")
            else:
                ws_abc.cell(row=row, column=12, value="Z")
        # Добавляем дашборд: сводная таблица ABC/XYZ
        start_summary_row = len(results) + 4
        ws_abc.cell(row=start_summary_row, column=1, value="Сводка ABC/XYZ")
        ws_abc.cell(row=start_summary_row, column=1).font = Font(bold=True, size=14)
        summary_data = [
            ["Класс", "Кол-во", "Средняя прибыль", "Средняя выручка"],
            ["A", 0, 0, 0],
            ["B", 0, 0, 0],
            ["C", 0, 0, 0]
        ]
        # Подсчёт
        for i, result in enumerate(results):
            cls = ws_abc.cell(row=i+2, column=4).value
            if cls == "A":
                summary_data[1][1] += 1
                summary_data[1][2] += result.gross_profit
                summary_data[1][3] += result.selling_price
            elif cls == "B":
                summary_data[2][1] += 1
                summary_data[2][2] += result.gross_profit
                summary_data[2][3] += result.selling_price
            else:
                summary_data[3][1] += 1
                summary_data[3][2] += result.gross_profit
                summary_data[3][3] += result.selling_price
        for r, data in enumerate(summary_data, start=start_summary_row+1):
            for c, val in enumerate(data, 1):
                ws_abc.cell(row=r, column=c, value=val)
                if r == start_summary_row+1:
                    ws_abc.cell(row=r, column=c).font = Font(bold=True)
        # Диаграммы (гистограмма ABC)
        if len(results) > 1:
            chart1 = BarChart()
            chart1.title = "Распределение ABC классов"
            data_ref = Reference(ws_abc, min_col=4, min_row=2, max_row=len(results)+1)
            cats_ref = Reference(ws_abc, min_col=1, min_row=2, max_row=len(results)+1)
            chart1.add_data(data_ref, titles_from_data=True)
            chart1.set_categories(cats_ref)
            chart1.legend = None
            ws_abc.add_chart(chart1, "I1")
        # Сохраняем
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()


# ============================================================================
# БЛОК 10: ИНТЕРФЕЙС ПОЛЬЗОВАТЕЛЯ (STREAMLIT)
# ============================================================================

def init_session_state():
    if 'api_manager' not in st.session_state:
        st.session_state.api_manager = YandexMarketAPIManager()
    if 'calculator' not in st.session_state:
        st.session_state.calculator = YandexMarketCalculator(
            api_manager=st.session_state.api_manager,
            tax_system="УСН 6% (доходы)"
        )
    if 'results' not in st.session_state:
        st.session_state.results = []
    if 'products' not in st.session_state:
        st.session_state.products = []
    if 'tax_system' not in st.session_state:
        st.session_state.tax_system = "УСН 6% (доходы)"
    if 'current_section' not in st.session_state:
        st.session_state.current_section = 'main'
    if 'recommendations' not in st.session_state:
        st.session_state.recommendations = []
    if 'onboarding_done' not in st.session_state:
        st.session_state.onboarding_done = False
    if 'uploaded_df' not in st.session_state:
        st.session_state.uploaded_df = None
    if 'column_mapping' not in st.session_state:
        st.session_state.column_mapping = {}
    if 'mapping_saved' not in st.session_state:
        st.session_state.mapping_saved = False
    if 'preferred_source' not in st.session_state:
        st.session_state.preferred_source = 'combined'  # 'api', 'ai', 'combined'

def show_onboarding():
    with st.expander("🎓 Новичок? Начни здесь!", expanded=not st.session_state.get('onboarding_done', False)):
        st.markdown("""
        ### 🚀 Быстрый старт за 3 шага:
        1. **Загрузи данные** — CSV с товарами (Артикул, Бренд, Весогабариты, Категория, Цена)
        2. **Настрой тарифы** — загрузи свои категории с тарифами или используй API
        3. **Рассчитай** — получи полную юнит-экономику для всех товаров

        ### 📋 Формат загрузки товаров:
        - `artikul` — артикул товара
        - `brand` — бренд
        - `category` — категория (должна совпадать с загруженными тарифами)
        - `selling_price` — цена продажи
        - `cogs` — себестоимость
        - `weight_kg` — вес в кг
        - `length_cm`, `width_cm`, `height_cm` — габариты (опционально)
        """)
        if st.button("✅ Понятно, больше не показывать"):
            st.session_state.onboarding_done = True
            st.rerun()

def render_sidebar():
    with st.sidebar:
        st.markdown("""
        <div style='text-align: center; padding: 20px 15px; background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460); border-radius: 12px; margin-bottom: 25px;'>
            <h1 style='color: white; margin: 0; font-size: 1.5em;'>🚀 FBS PRO</h1>
            <p style='color: #a8a8d0; margin: 8px 0 0 0; font-size: 0.9em;'>Яндекс Маркет</p>
            <p style='color: #6666aa; margin: 5px 0 0 0; font-size: 0.7em;'>v9.0.0 | Живые формулы</p>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("### 🧭 Навигация")
        sections = {
            "🏠 Главная": "main",
            "📦 Загрузка товаров": "upload",
            "📋 Категории и тарифы": "categories",
            "🧮 Калькулятор": "calculator",
            "📊 Результаты": "results",
            "📥 Экспорт Excel": "export",
            "💡 Рекомендации": "recommendations",
            "⚙️ Настройки": "settings"
        }
        selected_section = st.radio("Выберите раздел:", list(sections.keys()), label_visibility="collapsed")
        st.session_state.current_section = sections[selected_section]
        st.markdown("---")
        st.markdown("### 📊 Статус системы")
        calculator = st.session_state.calculator
        st.markdown(f"**💰 Налог:** {st.session_state.tax_system.split()[0]}")
        if calculator.tariffs_source == 'api':
            st.success("🔌 Тарифы: API Яндекс Маркет")
        elif calculator.tariffs_source == 'deepseek':
            st.info("🤖 Тарифы: DeepSeek AI")
        elif calculator.tariffs_source == 'user_categories':
            st.info("📄 Тарифы: Пользовательские категории")
        elif calculator.tariffs_source == 'user':
            st.info("👤 Тарифы: Пользовательский ввод")
        else:
            st.warning("⚠️ Тарифы: Не загружены!")
        if st.session_state.results:
            st.success(f"✅ Рассчитано: {len(st.session_state.results)} товаров")
            profitable = len([r for r in st.session_state.results if r.gross_profit > 0])
            st.metric("Прибыльных", f"{profitable} из {len(st.session_state.results)}")
        else:
            st.info("ℹ️ Расчеты не выполнялись")
        st.markdown("---")
        st.markdown("### ⚡ Быстрые действия")
        if st.button("🔄 Обновить тарифы (комб.)", use_container_width=True):
            with st.spinner("Загрузка тарифов..."):
                calculator.refresh_tariffs(force=True, preferred_source='combined')
                st.success("✅ Тарифы обновлены!")
                st.rerun()
        if st.button("🗑️ Очистить результаты", use_container_width=True):
            st.session_state.results = []
            st.session_state.products = []
            st.session_state.recommendations = []
            st.success("✅ Результаты очищены!")
            st.rerun()

def render_upload():
    st.markdown("## 📦 Загрузка товаров")
    st.info("""
    ### 📋 Формат данных:
    | Колонка | Описание | Обязательно |
    |---------|----------|-------------|
    | **artikul** | Артикул товара | ✅ Да |
    | **brand** | Бренд | ❌ Нет |
    | **category** | Категория (из вашего справочника) | ✅ Да |
    | **selling_price** | Цена продажи, ₽ | ✅ Да |
    | **cogs** | Себестоимость, ₽ | ✅ Да |
    | **weight_kg** | Вес, кг | ✅ Да |
    | **length_cm** | Длина, см | ❌ Нет |
    | **width_cm** | Ширина, см | ❌ Нет |
    | **height_cm** | Высота, см | ❌ Нет |
    ### 📌 Пример данных:
    artikul,brand,category,selling_price,cogs,weight_kg,length_cm,width_cm,height_cm
SKU-001,Samsung,electronics,15000,8000,0.5,15,10,2
SKU-002,Nike,clothing,5000,2000,0.3,30,20,5
    """)
    uploaded_file = st.file_uploader("📁 Загрузите CSV с товарами", type=['csv'])
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file)
            st.success(f"✅ Загружено {len(df)} товаров")
            with st.expander("👁️ Превью данных", expanded=True):
                st.dataframe(df.head(10), use_container_width=True)
            # Определение маппинга колонок
            st.markdown("### 🧩 Настройка соответствия колонок")
            # Стандартные поля, которые ожидаются
            field_names = {
                'artikul': 'Артикул',
                'brand': 'Бренд',
                'category': 'Категория',
                'selling_price': 'Цена продажи',
                'cogs': 'Себестоимость',
                'weight_kg': 'Вес (кг)',
                'length_cm': 'Длина (см)',
                'width_cm': 'Ширина (см)',
                'height_cm': 'Высота (см)',
                'warehouse_distance_km': 'Расстояние до склада (км)',
                'daily_sales': 'Продажи в день',
                'stock_depth_days': 'Глубина запаса (дней)'
            }
            col_mapping = {}
            for field, label in field_names.items():
                options = [''] + list(df.columns)
                default = field if field in df.columns else ''
                # Пытаемся найти лучшее совпадение
                if default == '':
                    for col in df.columns:
                        if col.lower().replace(' ', '_') == field.lower():
                            default = col
                            break
                selected = st.selectbox(f"Колонка для '{label}'", options, index=options.index(default) if default in options else 0, key=f"map_{field}")
                if selected:
                    col_mapping[field] = selected
            # Сохранить маппинг
            if st.button("💾 Сохранить маппинг", use_container_width=True):
                st.session_state.column_mapping = col_mapping
                st.session_state.mapping_saved = True
                st.success("✅ Маппинг сохранён! Теперь можно перейти к расчёту.")
            # Статистика по категориям
            if 'category' in df.columns:
                st.markdown("### 📊 Распределение по категориям")
                category_counts = df['category'].value_counts().reset_index()
                category_counts.columns = ['Категория', 'Количество']
                st.dataframe(category_counts, use_container_width=True)
            if st.button("🚀 ПЕРЕЙТИ К РАСЧЁТУ", type="primary", use_container_width=True):
                st.session_state.uploaded_df = df
                st.session_state.current_section = 'calculator'
                st.rerun()
        except Exception as e:
            st.error(f"❌ Ошибка чтения файла: {e}")

def render_categories():
    st.markdown("## 📋 Категории и тарифы")
    st.info("""
    ### 📌 Загрузка категорий с тарифами
    Загрузите CSV файл со своими категориями и тарифами.
    Формат должен содержать колонки (названия можно настроить через маппинг):
    | Рекомендуемое имя | Описание |
    |-------------------|----------|
    | **category** | Название категории |
    | **commission_rate** | Комиссия (в долях, например 0.15 = 15%) |
    | **min_commission** | Минимальная комиссия, ₽ |
    | **last_mile_base** | Базовая стоимость Last Mile, ₽ |
    | **last_mile_per_kg** | Стоимость за кг, ₽ |
    | **acquiring_fee** | Эквайринг (в долях) |
    | **return_fee** | Возвраты (в долях) |
    | **penalty_rate** | Штраф за просрочку (в долях) |
    """)

    # Выбор источника тарифов
    st.markdown("### 🔄 Источник тарифов")
    source_options = {
        "Комбинированный (API → AI)": "combined",
        "Только API Яндекс Маркет": "api",
        "Только DeepSeek AI": "ai"
    }
    selected_label = st.radio(
        "Выберите способ получения тарифов:",
        list(source_options.keys()),
        index=0
    )
    preferred_source = source_options[selected_label]
    st.session_state.preferred_source = preferred_source

    # Кнопка обновления тарифов с выбранным источником
    if st.button("🔄 Обновить тарифы", type="primary", use_container_width=True):
        with st.spinner(f"Загрузка тарифов из источника: {selected_label}..."):
            calculator = st.session_state.calculator
            calculator.refresh_tariffs(force=True, preferred_source=preferred_source)
            st.success(f"✅ Тарифы обновлены! Источник: {calculator.tariffs_source}")
            st.rerun()

    # Загрузка пользовательских категорий через CSV
    st.markdown("---")
    st.markdown("### 📂 Загрузка своих категорий (CSV)")
    uploaded_categories = st.file_uploader("📁 Загрузите CSV с категориями и тарифами", type=['csv'], key="categories_upload")
    if uploaded_categories:
        try:
            df = pd.read_csv(uploaded_categories)
            st.success(f"✅ Загружено {len(df)} категорий")
            with st.expander("👁️ Превью категорий", expanded=True):
                st.dataframe(df, use_container_width=True)
            # Маппинг для категорий
            st.markdown("### 🧩 Настройка соответствия колонок для тарифов")
            tariff_fields = {
                'category': 'Категория',
                'commission_rate': 'Комиссия (доля)',
                'min_commission': 'Мин. комиссия',
                'last_mile_base': 'Last Mile база',
                'last_mile_per_kg': 'Last Mile за кг',
                'acquiring_fee': 'Эквайринг (доля)',
                'return_fee': 'Возвраты (доля)',
                'penalty_rate': 'Штрафы (доля)'
            }
            tariff_mapping = {}
            for field, label in tariff_fields.items():
                options = [''] + list(df.columns)
                default = field if field in df.columns else ''
                if default == '':
                    for col in df.columns:
                        if col.lower().replace(' ', '_') == field.lower():
                            default = col
                            break
                selected = st.selectbox(f"Колонка для '{label}'", options, index=options.index(default) if default in options else 0, key=f"map_tariff_{field}")
                if selected:
                    tariff_mapping[field] = selected
            if st.button("📥 ЗАГРУЗИТЬ КАТЕГОРИИ В КАЛЬКУЛЯТОР", use_container_width=True):
                csv_content = uploaded_categories.getvalue().decode('utf-8')
                calculator = st.session_state.calculator
                calculator.refresh_tariffs(force=True, user_categories_csv=csv_content, column_mapping=tariff_mapping, preferred_source='combined')
                st.success(f"✅ Загружено {len(df)} категорий с тарифами!")
                st.rerun()
        except Exception as e:
            st.error(f"❌ Ошибка чтения файла: {e}")

    st.markdown("---")
    st.markdown("### 📊 Текущие тарифы")
    calculator = st.session_state.calculator
    if calculator.current_tariffs:
        df_tariffs = pd.DataFrame([
            {
                'Категория': cat,
                'Комиссия, %': round(data.get('commission_rate', 0) * 100, 2),
                'Мин. комиссия, ₽': data.get('min_commission', 0),
                'Last Mile база, ₽': data.get('last_mile_base', 0),
                'Last Mile за кг, ₽': data.get('last_mile_per_kg', 0),
                'Источник': data.get('source', 'unknown')
            }
            for cat, data in calculator.current_tariffs.items()
        ])
        st.dataframe(df_tariffs, use_container_width=True, height=300)
        st.caption(f"📊 Всего категорий: {len(df_tariffs)} | Источник: {calculator.tariffs_source} | Обновлено: {calculator.tariffs_updated_at.strftime('%d.%m.%Y %H:%M') if calculator.tariffs_updated_at else 'неизвестно'}")
    else:
        st.warning("⚠️ Нет загруженных тарифов. Загрузите категории или используйте API/AI.")

def render_calculator():
    st.markdown("## 🧮 Калькулятор FBS")
    if 'uploaded_df' not in st.session_state or st.session_state.uploaded_df is None:
        st.info("ℹ️ Сначала загрузите товары в разделе 'Загрузка товаров'")
        if st.button("📦 Перейти к загрузке", use_container_width=True):
            st.session_state.current_section = 'upload'
            st.rerun()
        return
    df = st.session_state.uploaded_df
    calculator = st.session_state.calculator
    st.success(f"✅ Загружено {len(df)} товаров")
    col1, col2 = st.columns(2)
    with col1:
        all_categories = df['category'].unique().tolist() if 'category' in df.columns else []
        selected_categories = st.multiselect("📂 Категории для расчёта", all_categories, default=all_categories)
    with col2:
        st.markdown("### ⚙️ Параметры расчёта")
        tax_system = st.selectbox("Система налогообложения", list(TAX_SYSTEMS.keys()),
                                 index=list(TAX_SYSTEMS.keys()).index(st.session_state.tax_system))
        if tax_system != st.session_state.tax_system:
            st.session_state.tax_system = tax_system
            calculator.tax_system = tax_system
    if selected_categories:
        df_filtered = df[df['category'].isin(selected_categories)]
    else:
        df_filtered = df
    st.markdown(f"📊 Будет рассчитано: **{len(df_filtered)}** товаров")
    if st.button("🚀 РАССЧИТАТЬ ВСЕ ТОВАРЫ", type="primary", use_container_width=True):
        with st.spinner("Выполняется расчёт..."):
            products = []
            mapping = st.session_state.column_mapping
            # Применяем маппинг для создания ProductData
            for _, row in df_filtered.iterrows():
                try:
                    # Получаем значения из колонок согласно маппингу
                    def get_val(field, default=0):
                        col = mapping.get(field)
                        if col and col in row:
                            return row[col]
                        # fallback на прямое имя
                        if field in row:
                            return row[field]
                        return default
                    product = ProductData(
                        artikul=str(get_val('artikul', '')),
                        brand=str(get_val('brand', '')),
                        category=str(get_val('category', 'default')),
                        selling_price=float(get_val('selling_price', 0)),
                        cogs=float(get_val('cogs', 0)),
                        weight_kg=float(get_val('weight_kg', 0)),
                        length_cm=float(get_val('length_cm', 0)),
                        width_cm=float(get_val('width_cm', 0)),
                        height_cm=float(get_val('height_cm', 0)),
                        warehouse_distance_km=float(get_val('warehouse_distance_km', 0)),
                        daily_sales=int(get_val('daily_sales', 5)),
                        stock_depth_days=int(get_val('stock_depth_days', 30))
                    )
                    products.append(product)
                except Exception as e:
                    st.warning(f"⚠️ Ошибка в строке {row.get('artikul', 'unknown')}: {e}")
            if products:
                results = calculator.calculate_batch(products)
                st.session_state.results = results
                st.session_state.products = products
                st.success(f"✅ Рассчитано {len(results)} товаров!")
                st.rerun()

def render_results():
    st.markdown("## 📊 Результаты расчётов")
    if not st.session_state.results:
        st.info("ℹ️ Нет результатов. Выполните расчёт в разделе 'Калькулятор'.")
        return
    results = st.session_state.results
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📦 Товаров", len(results))
    with col2:
        profitable = len([r for r in results if r.gross_profit > 0])
        st.metric("✅ Прибыльных", f"{profitable} ({profitable/len(results)*100:.0f}%)")
    with col3:
        total_profit = sum(r.gross_profit for r in results)
        st.metric("💰 Общая прибыль", f"{total_profit:,.0f} ₽")
    with col4:
        avg_margin = np.mean([r.margin_percent for r in results])
        st.metric("📊 Средняя маржа", f"{avg_margin:.1f}%")
    st.markdown("---")
    df_results = pd.DataFrame([r.get_summary() for r in results])
    st.dataframe(df_results, use_container_width=True, height=400)
    st.markdown("### 📈 Визуализация")
    col1, col2 = st.columns(2)
    with col1:
        margins = [r.margin_percent for r in results]
        fig = px.histogram(margins, title="Распределение маржинальности", labels={'value': 'Маржа, %', 'count': 'Количество'}, nbins=20)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        top = sorted(results, key=lambda x: x.gross_profit, reverse=True)[:10]
        fig = px.bar(x=[r.artikul[:15] for r in top], y=[r.gross_profit for r in top],
                    title="Топ-10 по прибыли", labels={'x': 'Артикул', 'y': 'Прибыль, ₽'})
        st.plotly_chart(fig, use_container_width=True)

def render_export():
    st.markdown("## 📥 Экспорт в Excel с живыми формулами")
    if not st.session_state.results:
        st.warning("⚠️ Нет данных для экспорта.")
        return
    if not OPENPYXL_AVAILABLE:
        st.error("❌ OpenPyXL не установлен. Установите: `pip install openpyxl`")
        return
    results = st.session_state.results
    products = st.session_state.products
    calculator = st.session_state.calculator
    st.info("""
    ### 📌 Особенности экспорта:
    1. **Живые формулы** — все расчёты пересчитываются при изменении данных
    2. **Условное форматирование** — прибыль/убыток выделены цветом
    3. **Цветовая шкала** — визуализация маржинальности
    4. **Итоговые строки** — сумма по всем товарам
    5. **Три листа**: Тарифы, Расчёты, ABC/XYZ с дашбордами
    """)
    st.success(f"✅ Доступно для экспорта: {len(results)} товаров")
    if st.button("📥 СКАЧАТЬ EXCEL С ФОРМУЛАМИ", type="primary", use_container_width=True):
        try:
            excel_bytes = ExcelExporter.export_with_formulas(
                results=results,
                products=products,
                tariffs=calculator.current_tariffs,
                tax_system=calculator.tax_system
            )
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"FBS_Yandex_Market_{timestamp}.xlsx"
            st.download_button(
                label="⬇️ Скачать Excel",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.success("✅ Excel файл с живыми формулами создан!")
        except Exception as e:
            st.error(f"❌ Ошибка создания Excel: {e}")

def render_recommendations():
    st.markdown("## 💡 Рекомендации")
    if not st.session_state.results:
        st.warning("⚠️ Нет данных. Выполните расчёт.")
        return
    calculator = st.session_state.calculator
    if st.button("🔄 Сгенерировать рекомендации", type="primary", use_container_width=True):
        with st.spinner("Генерация рекомендаций..."):
            st.session_state.recommendations = calculator.generate_recommendations(st.session_state.results)
        st.rerun()
    if st.session_state.recommendations:
        for rec in st.session_state.recommendations:
            priority_icon = "🔴" if rec['priority'] == 'high' else "🟡" if rec['priority'] == 'medium' else "🟢"
            with st.expander(f"{priority_icon} {rec['category']} - {rec['message'][:80]}..."):
                st.markdown(f"**{rec['message']}**")
                if rec.get('affected_products'):
                    st.markdown("**📦 Затронутые товары:**")
                    st.write(", ".join(rec['affected_products'][:10]))
    else:
        st.success("✅ Все показатели в норме! Рекомендаций нет.")

def render_settings():
    st.markdown("## ⚙️ Настройки")
    tab1, tab2 = st.tabs(["🔑 API Ключи", "🏛️ Налоговая система"])
    with tab1:
        st.markdown("### 🔑 API ключи Яндекс Маркет")
        st.info("API ключи используются для загрузки актуальных тарифов. Без них используются примерные значения.")
        api_manager = st.session_state.api_manager
        col1, col2 = st.columns(2)
        with col1:
            ym_token = st.text_input("Яндекс Маркет OAuth Token", 
                                    value=api_manager.get_api_key('yandex_market') or '', 
                                    type="password")
        with col2:
            ym_campaign = st.text_input("Campaign ID", 
                                       value=api_manager.get_api_key('yandex_campaign_id') or '')
        if st.button("💾 Сохранить Яндекс Маркет", use_container_width=True):
            if ym_token and ym_campaign:
                api_manager.save_api_key('yandex_market', ym_token)
                api_manager.save_api_key('yandex_campaign_id', ym_campaign)
                st.success("✅ Ключи Яндекс Маркет сохранены!")
        st.markdown("---")
        st.markdown("### 🤖 DeepSeek AI")
        ds_key = st.text_input("DeepSeek API Key", 
                              value=api_manager.get_api_key('deepseek') or '', 
                              type="password")
        if st.button("💾 Сохранить DeepSeek", use_container_width=True):
            if ds_key:
                api_manager.save_api_key('deepseek', ds_key)
                st.success("✅ Ключ DeepSeek сохранен!")
    with tab2:
        st.markdown("### 🏛️ Система налогообложения")
        tax_system = st.selectbox("Выберите систему", list(TAX_SYSTEMS.keys()),
                                 index=list(TAX_SYSTEMS.keys()).index(st.session_state.tax_system))
        if tax_system != st.session_state.tax_system:
            st.session_state.tax_system = tax_system
            st.session_state.calculator.tax_system = tax_system
            st.success(f"✅ Выбрано: {tax_system}")
        tax_config = TAX_SYSTEMS.get(tax_system, {})
        st.info(f"""
        **{tax_system}**
        - Ставка: {tax_config.get('rate', 0) * 100:.0f}%
        - База: {'Доходы' if tax_config.get('base') == 'revenue' else 'Прибыль'}
        """)


# ============================================================================
# БЛОК 11: ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

def main():
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="🚀",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    init_session_state()
    render_sidebar()
    current_section = st.session_state.get('current_section', 'main')
    calculator = st.session_state.calculator
    if current_section == 'main':
        st.markdown("""
        <div style='text-align: center; padding: 50px 30px; background: linear-gradient(135deg, #0f0c29, #302b63, #24243e); border-radius: 20px; margin-bottom: 35px;'>
            <h1 style='color: white; font-size: 3em; margin: 0;'>🚀 FBS Юнит-экономика</h1>
            <p style='color: #a8a8d0; font-size: 1.3em; margin: 20px 0;'>
                Яндекс Маркет — Живые формулы в Excel
            </p>
            <p style='color: #6666aa; font-size: 1em; margin: 10px 0;'>
                Загрузка товаров • Пользовательские категории • Расчёт 500K+ товаров
            </p>
        </div>
        """, unsafe_allow_html=True)
        show_onboarding()
        st.info("""
        ### 🎯 Что вы можете делать:
        | Раздел | Описание |
        |--------|----------|
        | 📦 **Загрузка товаров** | CSV с артикулами, брендами, категориями, ценами и весогабаритами |
        | 📋 **Категории и тарифы** | Загрузка своих категорий с тарифами, выбор источника (API, AI, комб.) |
        | 🧮 **Калькулятор** | Расчёт юнит-экономики для всех товаров |
        | 📊 **Результаты** | Просмотр и анализ расчётов |
        | 📥 **Экспорт Excel** | Скачивание Excel с живыми формулами и 3 листами |
        | 💡 **Рекомендации** | Автоматические рекомендации по оптимизации |
        ### 📌 Ключевые принципы:
        1. **НИКАКИХ ЗАХАРДКОЖЕННЫХ ТАРИФОВ** — все из API, AI, CSV
        2. **ЖИВЫЕ ФОРМУЛЫ** — при выгрузке в Excel
        3. **ПОЛЬЗОВАТЕЛЬСКИЕ КАТЕГОРИИ** — вы управляете тарифами
        """)
        if st.session_state.results:
            st.markdown("---")
            st.markdown("### 📊 Последние результаты")
            results = st.session_state.results[-5:]
            for r in results:
                color = "🟢" if r.gross_profit > 0 else "🔴"
                st.markdown(f"{color} **{r.artikul}** — Прибыль: {r.gross_profit:,.0f} ₽, Маржа: {r.margin_percent:.1f}%")
    elif current_section == 'upload':
        render_upload()
    elif current_section == 'categories':
        render_categories()
    elif current_section == 'calculator':
        render_calculator()
    elif current_section == 'results':
        render_results()
    elif current_section == 'export':
        render_export()
    elif current_section == 'recommendations':
        render_recommendations()
    elif current_section == 'settings':
        render_settings()
    st.markdown("---")
    st.caption(f"🚀 FBS Unit Economics PRO v{APP_VERSION} | Яндекс Маркет | "
              f"Источник тарифов: {calculator.tariffs_source.upper() if calculator.tariffs_source else 'НЕТ ДАННЫХ'} | "
              f"Данные актуальны на {datetime.now().strftime('%d.%m.%Y')}")

if __name__ == "__main__":
    main()
