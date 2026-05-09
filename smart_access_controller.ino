#include <SPI.h>
#include <MFRC522.h>
#include <Servo.h>
#include <EEPROM.h>

/*
  Smart Access Controller
  - Один скетч для:
    1) открытия двери по UID
    2) добавления/удаления UID в EEPROM
    3) записи ФИО и кабинета на карту по команде из Excel
    4) стирания данных карты и удаления UID из EEPROM
    5) отправки событий в Excel по Serial

  Протокол Serial (одна строка = одно событие):
    LOG|EVENT|UID|SURNAME|NAME|OFFICE|SOURCE|NOTE

  Команды из Excel:
    WRITE|SURNAME|NAME|OFFICE
    OPEN
    MODE|ACCESS
    MODE|UID
    MODE|WRITE
    MODE|ERASE
    CLEAR_UIDS
    STATUS
    PING

  Кнопки:
    MODE   - короткое нажатие: следующий режим
             долгое (>2 c): вернуть ACCESS
    ACTION - из любого режима: открыть дверь вручную и вернуться в ACCESS

  Режимы:
    ACCESS     - обычный проход по карте
    UID_MANAGE - скан карты добавляет/удаляет UID из EEPROM

  Сброс EEPROM:
    удерживать MODE + ACTION при старте > 3 секунд
*/

// ---------- Пины ----------
const byte PIN_SERVO      = 2;
const byte PIN_BUZZER     = 3;
const byte PIN_LED_RED    = 4;
const byte PIN_LED_GREEN  = 5;
const byte PIN_RFID_RST   = 6;
const byte PIN_RFID_SS    = 7;
const byte PIN_BTN_MODE   = 8;
const byte PIN_BTN_ACTION = 9;
const byte PIN_LED_BLUE   = 10;

// ---------- Настройки ----------
const unsigned long SERIAL_BAUD = 115200UL;
const unsigned long DOOR_OPEN_MS = 2000UL;
const unsigned long LONG_PRESS_MS = 2000UL;
const unsigned long STARTUP_RESET_MS = 3000UL;
const unsigned long DEBOUNCE_MS = 180UL;
const unsigned long LED_EVENT_MS = 2000UL;
const byte SERVO_OPEN_ANGLE = 110;
const byte SERVO_CLOSED_ANGLE = 0;
const byte RFID_BLOCK_SURNAME = 4;
const byte RFID_BLOCK_NAME    = 5;
const byte RFID_BLOCK_OFFICE  = 6;
const byte MAX_UID_RECORDS    = 20;
const byte UID_SLOT_SIZE      = 8;   // size(1) + up to 7 bytes UID, enough for common cards
const int EEPROM_MAGIC_ADDR   = 0;
const int EEPROM_COUNT_ADDR   = 1;
const byte EEPROM_MAGIC_VALUE = 0x5A;
const int EEPROM_UID_START    = 2;

MFRC522 rfid(PIN_RFID_SS, PIN_RFID_RST);
Servo lockServo;

enum WorkMode {
  MODE_ACCESS = 0,
  MODE_UID_MANAGE = 1,
  MODE_CARD_WRITE = 2,
  MODE_CARD_ERASE = 3
};

WorkMode currentMode = MODE_ACCESS;

struct PendingCardData {
  char surname[17];
  char name[17];
  char office[17];
  bool ready;
};

PendingCardData pendingWrite = {{0}, {0}, {0}, false};

bool doorOpen = false;
unsigned long doorOpenedAt = 0;
unsigned long lastModePressAt = 0;
unsigned long lastActionPressAt = 0;
bool eventLedActive = false;
unsigned long eventLedUntil = 0;

// -------------------- Вспомогательные функции --------------------

bool buttonPressed(byte pin) {
  return digitalRead(pin) == LOW;
}

void setLeds(bool greenOn, bool redOn) {
  digitalWrite(PIN_LED_GREEN, greenOn ? HIGH : LOW);
  digitalWrite(PIN_LED_RED, redOn ? HIGH : LOW);
}

void setBlueLed(bool blueOn) {
  digitalWrite(PIN_LED_BLUE, blueOn ? HIGH : LOW);
}

void startEventLedTimer() {
  eventLedActive = true;
  eventLedUntil = millis() + LED_EVENT_MS;
}

void beep(unsigned int freq, unsigned int durationMs, unsigned int pauseMs = 60) {
  tone(PIN_BUZZER, freq, durationMs);
  delay(durationMs + pauseMs);
  noTone(PIN_BUZZER);
}

void signalSuccess() {
  setBlueLed(false);
  setLeds(true, false);
  beep(1700, 120);
  startEventLedTimer();
}

void signalDenied() {
  setBlueLed(false);
  setLeds(false, true);
  beep(300, 140);
  beep(250, 160);
  startEventLedTimer();
}

void signalMode(byte count) {
  setBlueLed(false);
  for (byte i = 0; i < count; i++) {
    setLeds(true, true);
    beep(1200, 90);
  }
  setModeIndication();
}

void setModeIndication() {
  switch (currentMode) {
    case MODE_ACCESS:
      setLeds(false, false);
      setBlueLed(false);
      break;
    case MODE_UID_MANAGE:
      setLeds(false, false);
      setBlueLed(true);
      break;
    case MODE_CARD_WRITE:
      setLeds(false, false);
      setBlueLed(true);
      break;
    case MODE_CARD_ERASE:
      setLeds(false, false);
      setBlueLed(true);
      break;
  }
}

void handleEventLedTimeout() {
  if (eventLedActive && (long)(millis() - eventLedUntil) >= 0) {
    eventLedActive = false;
    setModeIndication();
  }
}

void logLine(const char* eventType, const String& uid, const String& surname,
             const String& name, const String& office,
             const char* source, const String& note) {
  Serial.print(F("LOG|"));
  Serial.print(eventType);
  Serial.print('|');
  Serial.print(uid);
  Serial.print('|');
  Serial.print(surname);
  Serial.print('|');
  Serial.print(name);
  Serial.print('|');
  Serial.print(office);
  Serial.print('|');
  Serial.print(source);
  Serial.print('|');
  Serial.println(note);
}

void sendStatus(const char* kind, const char* value) {
  Serial.print(F("STATUS|"));
  Serial.print(kind);
  Serial.print('|');
  Serial.println(value);
}

void sendUidMemoryStatus() {
  sendStatus("UID_COUNT", String(getUidCount()).c_str());
  sendStatus("UID_CAPACITY", String(MAX_UID_RECORDS).c_str());
}

void sendModeStatus() {
  switch (currentMode) {
    case MODE_ACCESS:
      sendStatus("MODE", "ACCESS");
      break;
    case MODE_UID_MANAGE:
      sendStatus("MODE", "UID_MANAGE");
      break;
    case MODE_CARD_WRITE:
      sendStatus("MODE", "CARD_WRITE");
      break;
    case MODE_CARD_ERASE:
      sendStatus("MODE", "CARD_ERASE");
      break;
  }
}

String uidToString(const MFRC522::Uid &uid) {
  String out;
  for (byte i = 0; i < uid.size; i++) {
    if (i) out += ' ';
    if (uid.uidByte[i] < 0x10) out += '0';
    out += String(uid.uidByte[i], HEX);
  }
  out.toUpperCase();
  return out;
}

void copyUidToSlot(byte* slot, const MFRC522::Uid &uid) {
  memset(slot, 0, UID_SLOT_SIZE);
  slot[0] = uid.size;
  for (byte i = 0; i < uid.size && i < UID_SLOT_SIZE - 1; i++) {
    slot[i + 1] = uid.uidByte[i];
  }
}

bool slotEqualsUid(const byte* slot, const MFRC522::Uid &uid) {
  if (slot[0] != uid.size) return false;
  for (byte i = 0; i < uid.size && i < UID_SLOT_SIZE - 1; i++) {
    if (slot[i + 1] != uid.uidByte[i]) return false;
  }
  return true;
}

byte getUidCount() {
  return EEPROM.read(EEPROM_COUNT_ADDR);
}

void setUidCount(byte value) {
  EEPROM.write(EEPROM_COUNT_ADDR, value);
}

int uidSlotAddress(byte index) {
  return EEPROM_UID_START + index * UID_SLOT_SIZE;
}

int findUidSlot(const MFRC522::Uid &uid) {
  byte count = getUidCount();
  byte slot[UID_SLOT_SIZE];

  for (byte i = 0; i < count; i++) {
    int addr = uidSlotAddress(i);
    for (byte j = 0; j < UID_SLOT_SIZE; j++) {
      slot[j] = EEPROM.read(addr + j);
    }
    if (slotEqualsUid(slot, uid)) return i;
  }
  return -1;
}

bool addUid(const MFRC522::Uid &uid) {
  byte count = getUidCount();
  if (count >= MAX_UID_RECORDS) return false;

  byte slot[UID_SLOT_SIZE];
  copyUidToSlot(slot, uid);
  int addr = uidSlotAddress(count);
  for (byte i = 0; i < UID_SLOT_SIZE; i++) {
    EEPROM.write(addr + i, slot[i]);
  }
  setUidCount(count + 1);
  return true;
}

bool removeUidByIndex(byte index) {
  byte count = getUidCount();
  if (index >= count) return false;

  for (byte i = index; i < count - 1; i++) {
    for (byte j = 0; j < UID_SLOT_SIZE; j++) {
      EEPROM.write(uidSlotAddress(i) + j, EEPROM.read(uidSlotAddress(i + 1) + j));
    }
  }

  for (byte j = 0; j < UID_SLOT_SIZE; j++) {
    EEPROM.write(uidSlotAddress(count - 1) + j, 0);
  }

  setUidCount(count - 1);
  return true;
}

void clearAllUids() {
  for (int i = 0; i < EEPROM.length(); i++) EEPROM.write(i, 0);
  EEPROM.write(EEPROM_MAGIC_ADDR, EEPROM_MAGIC_VALUE);
  EEPROM.write(EEPROM_COUNT_ADDR, 0);
}

bool ensureEepromInitialized() {
  if (EEPROM.read(EEPROM_MAGIC_ADDR) != EEPROM_MAGIC_VALUE) {
    clearAllUids();
    return true;
  }
  return false;
}

void servoMove(byte angle) {
  lockServo.attach(PIN_SERVO);
  lockServo.write(angle);
  delay(500);
  lockServo.detach();
}

void openDoor() {
  servoMove(SERVO_OPEN_ANGLE);
  doorOpen = true;
  doorOpenedAt = millis();
}

void closeDoor() {
  servoMove(SERVO_CLOSED_ANGLE);
  doorOpen = false;
  setModeIndication();
}

void normalizeField(char* text) {
  int len = strlen(text);
  while (len > 0 && (text[len - 1] == ' ' || text[len - 1] == '\r' || text[len - 1] == '\n' || text[len - 1] == '\t')) {
    text[len - 1] = '\0';
    len--;
  }
  int start = 0;
  while (text[start] == ' ' || text[start] == '\r' || text[start] == '\n' || text[start] == '\t') {
    start++;
  }
  if (start > 0) {
    memmove(text, text + start, strlen(text + start) + 1);
  }
}

void prepare16(char* dest16, const char* source) {
  memset(dest16, ' ', 16);
  for (byte i = 0; i < 16 && source[i] != '\0'; i++) {
    char c = source[i];
    if ((byte)c < 32) c = ' ';
    dest16[i] = c;
  }
}

bool writeBlock16(byte block, const char* source) {
  MFRC522::MIFARE_Key key;
  for (byte i = 0; i < 6; i++) key.keyByte[i] = 0xFF;

  byte data[16];
  prepare16((char*)data, source);

  MFRC522::StatusCode status;
  status = rfid.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A, block, &key, &(rfid.uid));
  if (status != MFRC522::STATUS_OK) return false;

  status = rfid.MIFARE_Write(block, data, 16);
  return status == MFRC522::STATUS_OK;
}

String readBlock16(byte block) {
  MFRC522::MIFARE_Key key;
  for (byte i = 0; i < 6; i++) key.keyByte[i] = 0xFF;

  byte buffer[18];
  byte size = sizeof(buffer);

  if (rfid.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A, block, &key, &(rfid.uid)) != MFRC522::STATUS_OK) {
    return "";
  }
  if (rfid.MIFARE_Read(block, buffer, &size) != MFRC522::STATUS_OK) {
    return "";
  }

  char text[17];
  for (byte i = 0; i < 16; i++) {
    text[i] = ((char)buffer[i] >= 32 && (char)buffer[i] <= 126) || ((byte)buffer[i] >= 192) ? (char)buffer[i] : ' ';
  }
  text[16] = '\0';
  normalizeField(text);
  return String(text);
}

bool writePendingCardData() {
  if (!pendingWrite.ready) return false;

  bool ok1 = writeBlock16(RFID_BLOCK_SURNAME, pendingWrite.surname);
  bool ok2 = writeBlock16(RFID_BLOCK_NAME, pendingWrite.name);
  bool ok3 = writeBlock16(RFID_BLOCK_OFFICE, pendingWrite.office);

  return ok1 && ok2 && ok3;
}

bool eraseCardData() {
  bool ok1 = writeBlock16(RFID_BLOCK_SURNAME, "");
  bool ok2 = writeBlock16(RFID_BLOCK_NAME, "");
  bool ok3 = writeBlock16(RFID_BLOCK_OFFICE, "");

  return ok1 && ok2 && ok3;
}

void clearPendingWrite() {
  pendingWrite.surname[0] = '\0';
  pendingWrite.name[0] = '\0';
  pendingWrite.office[0] = '\0';
  pendingWrite.ready = false;
}

void setMode(WorkMode mode) {
  currentMode = mode;
  setModeIndication();
  sendModeStatus();

  switch (currentMode) {
    case MODE_ACCESS:
      signalMode(1);
      break;
    case MODE_UID_MANAGE:
      signalMode(2);
      break;
    case MODE_CARD_WRITE:
      signalMode(3);
      break;
    case MODE_CARD_ERASE:
      signalMode(4);
      break;
  }
}

void nextMode() {
  if (currentMode == MODE_ACCESS) setMode(MODE_UID_MANAGE);
  else setMode(MODE_ACCESS);
}

void processSerialCommand(String line) {
  line.trim();
  if (line.length() == 0) return;

  if (line == "PING") {
    sendStatus("PING", "OK");
    return;
  }

  if (line == "STATUS") {
    sendModeStatus();
    sendUidMemoryStatus();
    return;
  }

  if (line == "OPEN") {
    openDoor();
    signalSuccess();
    logLine("MANUAL_OPEN", "", "", "", "", "EXCEL", "Команда OPEN");
    return;
  }

  if (line == "CLEAR_UIDS") {
    clearAllUids();
    clearPendingWrite();
    setMode(MODE_ACCESS);
    signalSuccess();
    logLine("UIDS_CLEARED", "", "", "", "", "EXCEL", "Все UID удалены из памяти Arduino");
    sendUidMemoryStatus();
    return;
  }

  if (line == "MODE|ACCESS") {
    setMode(MODE_ACCESS);
    return;
  }
  if (line == "MODE|UID") {
    setMode(MODE_UID_MANAGE);
    return;
  }
  if (line == "MODE|WRITE") {
    setMode(MODE_CARD_WRITE);
    return;
  }
  if (line == "MODE|ERASE") {
    clearPendingWrite();
    setMode(MODE_CARD_ERASE);
    return;
  }

  if (line.startsWith("WRITE|")) {
    int p1 = line.indexOf('|');
    int p2 = line.indexOf('|', p1 + 1);
    int p3 = line.indexOf('|', p2 + 1);

    if (p2 < 0 || p3 < 0) {
      sendStatus("ERROR", "WRITE_FORMAT");
      return;
    }

    String s = line.substring(p1 + 1, p2);
    String n = line.substring(p2 + 1, p3);
    String o = line.substring(p3 + 1);

    s.trim(); n.trim(); o.trim();

    s.toCharArray(pendingWrite.surname, sizeof(pendingWrite.surname));
    n.toCharArray(pendingWrite.name, sizeof(pendingWrite.name));
    o.toCharArray(pendingWrite.office, sizeof(pendingWrite.office));
    normalizeField(pendingWrite.surname);
    normalizeField(pendingWrite.name);
    normalizeField(pendingWrite.office);
    pendingWrite.ready = true;

    setMode(MODE_CARD_WRITE);
    sendStatus("WRITE", "READY");
    return;
  }

  sendStatus("ERROR", "UNKNOWN_COMMAND");
}

void processSerialInput() {
  static String line;
  while (Serial.available()) {
    char c = (char)Serial.read();
    if (c == '\n') {
      processSerialCommand(line);
      line = "";
    } else if (c != '\r') {
      line += c;
    }
  }
}

void handleButtons() {
  unsigned long now = millis();

  if (buttonPressed(PIN_BTN_MODE)) {
    if (lastModePressAt == 0) lastModePressAt = now;
  } else if (lastModePressAt != 0) {
    unsigned long held = now - lastModePressAt;
    lastModePressAt = 0;
    if (held >= LONG_PRESS_MS) setMode(MODE_ACCESS);
    else nextMode();
  }

  if (buttonPressed(PIN_BTN_ACTION)) {
    if (lastActionPressAt == 0) lastActionPressAt = now;
  } else if (lastActionPressAt != 0) {
    unsigned long held = now - lastActionPressAt;
    lastActionPressAt = 0;
    if (held >= DEBOUNCE_MS) {
      if (currentMode != MODE_ACCESS) {
        currentMode = MODE_ACCESS;
        sendStatus("MODE", "ACCESS");
        clearPendingWrite();
      }
      openDoor();
      signalSuccess();
      logLine("MANUAL_OPEN", "", "", "", "", "BUTTON", "Открыто кнопкой ACTION, режим ACCESS");
    }
  }
}

void handleDoorTimeout() {
  if (doorOpen && millis() - doorOpenedAt >= DOOR_OPEN_MS) {
    closeDoor();
  }
}

void handleCardScan() {
  if (!rfid.PICC_IsNewCardPresent()) return;
  if (!rfid.PICC_ReadCardSerial()) return;

  String uid = uidToString(rfid.uid);

  if (currentMode == MODE_ACCESS) {
    int idx = findUidSlot(rfid.uid);
    if (idx >= 0) {
      String surname = readBlock16(RFID_BLOCK_SURNAME);
      String name = readBlock16(RFID_BLOCK_NAME);
      String office = readBlock16(RFID_BLOCK_OFFICE);

      openDoor();
      signalSuccess();
      logLine("ACCESS_GRANTED", uid, surname, name, office, "CARD", "");
    } else {
      signalDenied();
      logLine("ACCESS_DENIED", uid, "", "", "", "CARD", "UID не найден в EEPROM");
    }
  } else if (currentMode == MODE_UID_MANAGE) {
    int idx = findUidSlot(rfid.uid);
    if (idx >= 0) {
      removeUidByIndex((byte)idx);
      signalSuccess();
      logLine("UID_REMOVED", uid, "", "", "", "CARD", "Карта удалена из списка доступа");
      sendUidMemoryStatus();
    } else {
      if (addUid(rfid.uid)) {
        signalSuccess();
        logLine("UID_ADDED", uid, "", "", "", "CARD", "Карта добавлена в список доступа");
        sendUidMemoryStatus();
      } else {
        signalDenied();
        logLine("UID_ADD_FAILED", uid, "", "", "", "CARD", "EEPROM заполнена");
        sendUidMemoryStatus();
      }
    }
  } else if (currentMode == MODE_CARD_WRITE) {
    if (!pendingWrite.ready) {
      signalDenied();
      logLine("CARD_WRITE_SKIPPED", uid, "", "", "", "CARD", "Нет подготовленных данных WRITE|... из Excel");
    } else {
      if (writePendingCardData()) {
        signalSuccess();
        logLine("CARD_WRITTEN", uid, pendingWrite.surname, pendingWrite.name, pendingWrite.office, "EXCEL", "Данные записаны на карту");
        clearPendingWrite();
        setMode(MODE_ACCESS);
      } else {
        signalDenied();
        logLine("CARD_WRITE_FAILED", uid, pendingWrite.surname, pendingWrite.name, pendingWrite.office, "EXCEL", "Ошибка записи блоков 4/5/6");
      }
    }
  } else if (currentMode == MODE_CARD_ERASE) {
    bool dataErased = eraseCardData();
    int idx = findUidSlot(rfid.uid);
    bool uidRemoved = false;
    if (idx >= 0) {
      uidRemoved = removeUidByIndex((byte)idx);
    }

    if (dataErased && (idx < 0 || uidRemoved)) {
      signalSuccess();
      logLine("CARD_ERASED", uid, "", "", "", "EXCEL", idx >= 0 ? "Данные карты стерты, UID удален из списка доступа" : "Данные карты стерты, UID не был в списке доступа");
      sendUidMemoryStatus();
      setMode(MODE_ACCESS);
    } else {
      signalDenied();
      logLine("CARD_ERASE_FAILED", uid, "", "", "", "EXCEL", "Ошибка стирания карты или удаления UID");
    }
  }

  rfid.PICC_HaltA();
  rfid.PCD_StopCrypto1();
  delay(250);
}

void handleStartupReset() {
  unsigned long startedAt = millis();
  while (buttonPressed(PIN_BTN_MODE) && buttonPressed(PIN_BTN_ACTION)) {
    if (millis() - startedAt >= STARTUP_RESET_MS) {
      clearAllUids();
      for (byte i = 0; i < 3; i++) {
        setLeds(true, true);
        beep(900, 120);
      }
      logLine("EEPROM_CLEARED", "", "", "", "", "SYSTEM", "MODE+ACTION удерживались при старте");
      break;
    }
  }
}

// -------------------- setup / loop --------------------

void setup() {
  pinMode(PIN_BTN_MODE, INPUT_PULLUP);
  pinMode(PIN_BTN_ACTION, INPUT_PULLUP);
  pinMode(PIN_LED_RED, OUTPUT);
  pinMode(PIN_LED_GREEN, OUTPUT);
  pinMode(PIN_LED_BLUE, OUTPUT);
  pinMode(PIN_BUZZER, OUTPUT);

  Serial.begin(SERIAL_BAUD);
  SPI.begin();
  rfid.PCD_Init();

  ensureEepromInitialized();
  handleStartupReset();
  closeDoor();
  setMode(MODE_ACCESS);

  sendStatus("BOOT", "READY");
  sendUidMemoryStatus();
}

void loop() {
  processSerialInput();
  handleButtons();
  handleDoorTimeout();
  handleEventLedTimeout();
  handleCardScan();
}
