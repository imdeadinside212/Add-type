# Add type -- Auto Build + Release + Update (mien phi 100%)

Repo: https://github.com/imdeadinside212/add-type

Bo file nay giup app "Add type" (Python, build bang PyInstaller) tu dong:
1. **Build** ban Windows va macOS moi khi push 1 tag version.
2. **Release** len GitHub Releases (mien phi voi repo public).
3. **Tu update** o phia nguoi dung khi mo app -- co thanh trang thai nho
   o goc duoi cua so va nut "Update" khi co ban moi.

## Cau truc file trong repo

```
add-type/
├── main.py                          <- code nguon chinh, dung chung cho ca Windows va Mac
│                                        (tu re nhanh theo sys.platform ben trong)
├── icon.ico                          <- icon cho ban Windows (ban tu them vao)
├── icon.icns                         <- (tuy chon) icon rieng cho ban Mac, khong co thi build van chay
├── updater.py                        <- module tu-update, da dien san REPO = "imdeadinside212/add-type"
├── main_example.py                   <- file tham khao cach goi updater.py (KHONG bat buoc,
│                                        code that da duoc gop thang vao main.py roi)
├── requirements.txt                  <- danh sach thu vien can khi BUILD (khong anh huong nguoi dung cuoi)
└── .github/
    └── workflows/
        └── build-release.yml         <- workflow build + release tu dong
```

## Trang thai hien tai

- `main.py`: da gop fix ket noi Photoshop tren Mac (chuyen tu JXA sang AppleScript,
  tu tat popup canh bao script cua Photoshop) + fix icon rieng cho macOS (.icns) +
  giu nguyen cac fix moi nhat ve xu ly layer trong Photoshop (khong xu ly trung
  Smart Object).
- `build-release.yml`: ca 2 job Windows va Mac deu build tu **cung 1 file `main.py`**
  (khong can file rieng cho Mac nua).
- App co san thanh trang thai + nut "Update" o goc duoi cua so, tu kiem tra
  ban moi ngay luc mo app.

## Cach ra ban update moi (moi lan sau nay)

1. Sua code trong `main.py`.
2. **Tang so o dong `APP_VERSION = "..."` trong `main.py`** (vi du "1.0.0" -> "1.0.1")
   -- neu quen buoc nay, nguoi dang dung ban cu se KHONG thay nut Update xuat hien.
3. Commit + push code len nhanh `main` (qua GitHub Desktop hoac dong lenh git).
4. Tao tag & push tag (kich hoat build):
   ```bash
   git tag v1.0.1
   git push origin v1.0.1
   ```
   Hoac tren web: repo -> Releases -> "Draft a new release" -> go tag `v1.0.1`
   o o "Choose a tag" -> "Create new tag ... on publish" -> Publish release.
5. Doi GitHub Actions build xong (~3-5 phut, xem o tab Actions), sau do vao tab
   Releases se thay `Add type.exe` va `Add type-mac.zip` cua ban moi.
6. Nguoi dung dang chay ban cu, lan tiep theo mo app, se tu thay nut "Update"
   xuat hien o goc duoi -- bam vao la tu tai + cai + khoi dong lai ban moi.

## Nhung diem can luu y

- **Windows SmartScreen**: exe chua ky se bi canh bao "Windows protected your PC".
  Nguoi dung bam "More info" -> "Run anyway" la duoc. Het canh bao hoan toan can
  mua chung chi code-signing (khong free).
- **macOS Gatekeeper**: app chua notarize se bi chan "unidentified developer" o
  **lan cai dat dau tien**. Huong dan nguoi dung chuot phai -> Open thay vi double-click,
  hoac chay `xattr -cr` mot lan. Tu lan update tu dong thu 2 tro di, `updater.py` da
  tu chay san `xattr -cr` nen it bi canh bao hon.
- **requirements.txt**: chi dung luc GitHub Actions BUILD app, khong lien quan gi
  toi may nguoi dung cuoi -- ho chi tai file exe/app hoan chinh, khong can cai gi ca.
  `pywin32` da duoc danh dau chi cai tren Windows (`sys_platform == "win32"`) de
  khong lam hong buoc build tren may ao Mac.
- **icon.icns**: khong bat buoc -- neu chua co, build-release.yml se tu bo qua va
  build app Mac voi icon mac dinh, khong loi gi ca.
- **GitHub Actions quota**: mien phi khong gioi han voi repo public.
