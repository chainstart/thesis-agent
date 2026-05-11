from __future__ import annotations

import re


def normalize_reference_text(text: str) -> str:
    normalized = re.sub(r"\s+", " ", text).strip()
    if not normalized:
        return normalized
    normalized = normalized.replace("丶", " ")
    normalized = re.sub(r"^\[\d+\]\s*", "", normalized, count=1).strip()

    curated = _curated_reference(normalized)
    if curated is not None:
        return curated

    normalized = re.sub(r"https?://\S+", "", normalized)
    normalized = re.sub(r"\s+([,.;:])", r"\1", normalized)
    normalized = re.sub(r"(?<=[\u4e00-\u9fff\]\)])\.(?=[A-Za-z\u4e00-\u9fff\[])", ". ", normalized)
    normalized = re.sub(r"(?<=\])\.(?=\S)", ". ", normalized)
    normalized = re.sub(r",(?=\d{4})", ", ", normalized)
    normalized = re.sub(r"(?<=\d{4}),(?=[\d(（])", ", ", normalized)
    normalized = re.sub(r"(?<=[\u4e00-\u9fff]),(?=[\u4e00-\u9fff])", ", ", normalized)
    normalized = re.sub(r"\.(?=DOI:)", ". ", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"\s+([。；，、])", r"\1", normalized)
    normalized = re.sub(r"、\s+", "、", normalized)
    normalized = re.sub(r"([。；，])(?=\S)", r"\1 ", normalized)
    normalized = re.sub(r"\s{2,}", " ", normalized).strip(" .")
    return normalized + "." if normalized and not normalized.endswith((".", "。")) else normalized


def _curated_reference(text: str) -> str | None:
    compact = re.sub(r"\s+", "", text).lower()
    if "危险化学品安全管理条例" in compact:
        return "中华人民共和国国务院. 危险化学品安全管理条例: 中华人民共和国国务院令第591号[S]. 2011."
    if "dht11" in compact and ("dht22" in compact or "am2302" in compact):
        return "Aosong Electronics Co., Ltd. DHT11 and AM2302/DHT22 digital temperature and humidity sensor datasheets[Z]. Guangzhou: Aosong Electronics Co., Ltd., 2015."
    if "dht11" in compact:
        return "Aosong Electronics Co., Ltd. DHT11 temperature and humidity sensor datasheet[Z]. Guangzhou: Aosong Electronics Co., Ltd., 2015."
    if "dht22" in compact or "am2302" in compact:
        return "Aosong Electronics Co., Ltd. AM2302/DHT22 digital temperature and humidity sensor datasheet[Z]. Guangzhou: Aosong Electronics Co., Ltd., 2015."
    if "sgp30" in compact:
        return "Sensirion AG. SGP30 multi-pixel gas sensor datasheet[Z]. Stafa: Sensirion AG, 2020."
    if "sgp40" in compact:
        return "Sensirion AG. SGP40 VOC sensor datasheet[Z]. Stafa: Sensirion AG, 2021."
    if "bme680" in compact:
        return "Bosch Sensortec GmbH. BME680 low power gas, pressure, temperature and humidity sensor datasheet[Z]. Reutlingen: Bosch Sensortec GmbH, 2024."
    if "stm32f103x8" in compact or "stm32f103xb" in compact:
        return "STMicroelectronics. STM32F103x8, STM32F103xB medium-density performance line Arm-based 32-bit MCU datasheet[Z]. Geneva: STMicroelectronics, 2022."
    if "rm0008" in compact or "stm32f1系列微控制器参考手册" in compact or "stm32f1" in compact and "参考手册" in compact:
        return "STMicroelectronics. RM0008 reference manual: STM32F101xx, STM32F102xx, STM32F103xx, STM32F105xx and STM32F107xx advanced Arm-based 32-bit MCUs[Z]. Geneva: STMicroelectronics, 2024."
    if "mfrc522" in compact or "rc522" in compact:
        return "NXP Semiconductors. MFRC522 standard performance MIFARE and NTAG frontend datasheet[Z]. Eindhoven: NXP Semiconductors, 2016."
    if "mqttversion5.0" in compact:
        return "OASIS. MQTT Version 5.0[S]. 2019."
    if "mqttversion3.1.1" in compact:
        return "OASIS. MQTT Version 3.1.1[S]. 2014."
    if "quickintroductiontothemqttprotocol" in compact or "mqtt协议简介" in compact or "scaleway" in compact:
        return "OASIS. MQTT Version 5.0[S]. 2019."
    if "宗文锦" in compact and "危化品存储柜安全管控系统" in compact:
        return "宗文锦. 危化品存储柜安全管控系统研究与开发[D]. 无锡: 江南大学, 2021."
    return None
