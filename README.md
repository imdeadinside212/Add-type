# Add type -- Auto Build + Release + Update (mien phi 100%)


## Nhung diem can luu y

- **Windows SmartScreen**: exe chua ky se bi canh bao "Windows protected your PC".
  Nguoi dung bam "More info" -> "Run anyway" la duoc. Het canh bao hoan toan can
  mua chung chi code-signing (khong free).
- **macOS Gatekeeper**: app chua notarize se bi chan "unidentified developer" o
  **lan cai dat dau tien**. Huong dan nguoi dung chuot phai -> Open thay vi double-click,
  hoac chay `xattr -cr` mot lan. Tu lan update tu dong thu 2 tro di, `updater.py` da
  tu chay san `xattr -cr` nen it bi canh bao hon.
