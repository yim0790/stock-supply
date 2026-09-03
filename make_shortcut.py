# 바탕화면에 '재고수급 업데이트'(배치 실행)와 '판매재고수급 현황'(웹 조회) 바로가기 두 개를 만드는 스크립트
# -*- coding: utf-8 -*-
r"""
make_shortcut.bat 이 호출한다. 한글 이름·특수문자 경로를 다루므로 배치가 아닌 파이썬이 담당한다.
이름을 바꾸려면 아래 두 상수만 수정하면 된다.

만들어지는 것
  바탕화면\재고수급 업데이트.lnk      → update.bat 실행   (아이콘 icons\update icon.ico)
  바탕화면\판매재고수급 현황.url      → GitHub Pages 열기 (아이콘 icons\view icon.ico)
  직원배포용\판매재고수급 현황.url    → 직원에게 보낼 사본 (아이콘 없음 - 상대 PC엔 파일이 없으므로)
"""

import os, re, subprocess

MY_SHORTCUT_NAME    = "재고수급 업데이트"    # 내 PC 바탕화면 (update.bat 실행)
STAFF_SHORTCUT_NAME = "판매재고수급 현황"    # 조회용 웹 바로가기 (내 바탕화면 + 직원 배포용)

HERE = os.path.dirname(os.path.abspath(__file__))
UPDATE_ICO = os.path.join(HERE, "icons", "update icon.ico")
VIEW_ICO   = os.path.join(HERE, "icons", "view icon.ico")


def read_config():
    """config.cmd 에서 set "KEY=VALUE" 줄을 읽어온다."""
    cfg = {}
    path = os.path.join(HERE, "config.cmd")
    if not os.path.exists(path):
        raise SystemExit("[중단] config.cmd 가 없습니다.")
    with open(path, encoding="ascii", errors="ignore") as f:
        for line in f:
            m = re.match(r'\s*set\s+"([A-Z_]+)=(.*)"\s*$', line)
            if m:
                cfg[m.group(1)] = m.group(2).strip()
    return cfg


def desktop():
    r = subprocess.run(["powershell", "-NoProfile", "-Command",
                        "[Environment]::GetFolderPath('Desktop')"],
                       capture_output=True, text=True)
    return (r.stdout or "").strip()


def make_lnk(name, target, workdir, icon):
    """PowerShell COM 으로 바탕화면 .lnk 생성. 한글은 환경변수로 넘겨 인코딩 문제를 피한다."""
    env = dict(os.environ, SC_NAME=name, SC_TARGET=target, SC_WORKDIR=workdir, SC_ICON=icon)
    ps = (
        "$ws = New-Object -ComObject WScript.Shell; "
        "$p = [Environment]::GetFolderPath('Desktop') + '\\' + $env:SC_NAME + '.lnk'; "
        "$s = $ws.CreateShortcut($p); "
        "$s.TargetPath = $env:SC_TARGET; "
        "$s.WorkingDirectory = $env:SC_WORKDIR; "
        "$s.IconLocation = $env:SC_ICON; "
        "$s.Save(); "
        "Write-Host $p"
    )
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       env=env, capture_output=True, text=True)
    if r.returncode != 0:
        print("  [실패]", (r.stderr or "").strip()[:300])
        return None
    return (r.stdout or "").strip()


def write_url(path, site, icon=None):
    """인터넷 바로가기(.url)는 단순 텍스트다. 아이콘 경로에 한글이 있으므로 cp949 → utf-8 순으로 시도한다."""
    lines = ["[InternetShortcut]", f"URL={site}"]
    if icon:
        lines += [f"IconFile={icon}", "IconIndex=0"]
    else:
        lines += ["IconIndex=0"]
    body = "\n".join(lines) + "\n"
    for enc in ("cp949", "utf-8"):
        try:
            with open(path, "w", encoding=enc, newline="\r\n") as f:
                f.write(body)
            return path
        except UnicodeEncodeError:
            continue
    return None


def main():
    cfg = read_config()
    user, repo = cfg.get("GH_USER", ""), cfg.get("GH_REPO", "")
    if not user or user == "YOUR-GITHUB-ID":
        raise SystemExit("[중단] config.cmd 의 GH_USER / GH_REPO 를 먼저 채워주세요.")
    site = f"https://{user}.github.io/{repo}/"
    dsk = desktop()

    print("=" * 60)
    print("바탕화면 바로가기 만들기")
    print("=" * 60)

    # 1) 업데이트 버튼 (update.bat 실행)
    print(f"[1/3] 바탕화면 '{MY_SHORTCUT_NAME}'  (업데이트 실행)")
    if not os.path.exists(UPDATE_ICO):
        print(f"      [주의] 아이콘이 없습니다: {UPDATE_ICO} — 기본 아이콘으로 만듭니다")
    p = make_lnk(MY_SHORTCUT_NAME, os.path.join(HERE, "update.bat"), HERE,
                 UPDATE_ICO if os.path.exists(UPDATE_ICO) else "imageres.dll,109")
    print("      생성:", p or "실패")

    # 2) 조회 바로가기 (내 바탕화면 - 로컬 아이콘 사용)
    print(f"[2/3] 바탕화면 '{STAFF_SHORTCUT_NAME}'  (웹 조회)")
    if not dsk:
        print("      [실패] 바탕화면 경로를 찾지 못했습니다")
    else:
        p = write_url(os.path.join(dsk, f"{STAFF_SHORTCUT_NAME}.url"), site,
                      VIEW_ICO if os.path.exists(VIEW_ICO) else None)
        print("      생성:", p or "실패")

    # 3) 직원 배포용 (아이콘 경로는 넣지 않는다 - 상대 PC엔 그 파일이 없다)
    print(f"[3/3] 직원배포용 '{STAFF_SHORTCUT_NAME}.url'")
    outdir = os.path.join(HERE, "직원배포용")
    os.makedirs(outdir, exist_ok=True)
    p = write_url(os.path.join(outdir, f"{STAFF_SHORTCUT_NAME}.url"), site)
    print("      생성:", p or "실패")
    print("      주소:", site)
    print()
    print("바탕화면에 아이콘 2개가 생겼습니다.")
    print("  · 재고수급 업데이트  → 데이터 갱신 + 사이트 반영")
    print("  · 판매재고수급 현황  → 화면 열기")
    print()
    print("'직원배포용' 폴더의 .url 파일을 카톡/메일로 보내면,")
    print("직원은 바탕화면에 끌어다 놓기만 하면 됩니다.")


if __name__ == "__main__":
    main()
