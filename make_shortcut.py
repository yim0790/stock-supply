# 바탕화면 바로가기(재고수급 업데이트 버튼)와 직원 배포용 웹 바로가기(.url)를 만드는 스크립트
# -*- coding: utf-8 -*-
"""
make_shortcut.bat 이 호출한다. 한글 이름을 다루므로 배치가 아닌 파이썬이 담당한다.
바로가기 이름을 바꾸고 싶으면 아래 두 상수만 수정하면 된다.
"""

import os, re, subprocess

MY_SHORTCUT_NAME    = "재고수급 업데이트"   # 내 PC 바탕화면 (update.bat 실행)
STAFF_SHORTCUT_NAME = "판매재고수급 현황"   # 직원에게 나눠줄 웹 바로가기

HERE = os.path.dirname(os.path.abspath(__file__))


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


def make_lnk(name, target, workdir):
    """PowerShell COM 으로 바탕화면 .lnk 생성. 한글 이름은 환경변수로 전달한다."""
    env = dict(os.environ, SC_NAME=name, SC_TARGET=target, SC_WORKDIR=workdir)
    ps = (
        "$ws = New-Object -ComObject WScript.Shell; "
        "$p = [Environment]::GetFolderPath('Desktop') + '\\' + $env:SC_NAME + '.lnk'; "
        "$s = $ws.CreateShortcut($p); "
        "$s.TargetPath = $env:SC_TARGET; "
        "$s.WorkingDirectory = $env:SC_WORKDIR; "
        "$s.IconLocation = 'imageres.dll,109'; "
        "$s.Save(); "
        "Write-Host $p"
    )
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       env=env, capture_output=True, text=True)
    if r.returncode != 0:
        print("  [실패]", (r.stderr or "").strip()[:300])
        return None
    return (r.stdout or "").strip()


def main():
    cfg = read_config()
    user, repo = cfg.get("GH_USER", ""), cfg.get("GH_REPO", "")
    if not user or user == "YOUR-GITHUB-ID":
        raise SystemExit("[중단] config.cmd 의 GH_USER / GH_REPO 를 먼저 채워주세요.")
    site = f"https://{user}.github.io/{repo}/"

    print("=" * 60)
    print("바로가기 만들기")
    print("=" * 60)

    # 1) 사장님 PC: 업데이트 버튼
    print(f"[1/2] 내 바탕화면 바로가기 '{MY_SHORTCUT_NAME}'")
    p = make_lnk(MY_SHORTCUT_NAME, os.path.join(HERE, "update.bat"), HERE)
    if p:
        print("      생성:", p)

    # 2) 직원 배포용 웹 바로가기 (.url 은 단순 텍스트 파일이라 그냥 만들면 된다)
    print(f"[2/2] 직원 배포용 '{STAFF_SHORTCUT_NAME}.url'")
    outdir = os.path.join(HERE, "직원배포용")
    os.makedirs(outdir, exist_ok=True)
    urlfile = os.path.join(outdir, f"{STAFF_SHORTCUT_NAME}.url")
    with open(urlfile, "w", encoding="ascii", newline="\r\n") as f:
        f.write("[InternetShortcut]\n")
        f.write(f"URL={site}\n")
        f.write("IconIndex=0\n")
    print("      생성:", urlfile)
    print("      주소:", site)
    print()
    print("이 '직원배포용' 폴더의 파일을 카톡/메일로 직원에게 보내고,")
    print("직원은 바탕화면에 끌어다 놓기만 하면 됩니다.")


if __name__ == "__main__":
    main()
