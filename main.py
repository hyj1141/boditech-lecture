import win32com.client
import datetime
import os


def download_and_merge_ppts():
    # 1. 파일들을 저장할 폴더 생성
    save_folder = os.path.join(os.getcwd(), "오늘의_회의자료")
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)

    # 🎯 [핵심] 찾고자 하는 파일명 속 날짜 키워드 생성 (예: "260521")
    today = datetime.date.today()
    target_date_str = today.strftime("%y%m%d")

    expected_teams = {
        "허용민": (1, "연구기획팀"),
        "김태원": (2, "심혈관팀"),
        "한예지": (3, "급성감염팀"),
        "정진용": (4, "Cancer팀"),
        "이소희": (5, "호르몬팀"),
        "김영은": (6, "치료용항체팀"),
        "김세희": (7, "갑상선팀"),
        "함은선": (8, "당뇨팀")
    }

    downloaded_files = []
    received_senders = set()  # '260521'이 적힌 정상 PPT를 보낸 사람
    wrong_date_senders = set()  # PPT는 보냈으나 파일명에 '260521'이 없는 사람 (예: 260514 등)

    try:
        # ==========================================
        # [1단계] 엄격한 파일명 선별 및 다운로드
        # ==========================================
        print(f"1. 파일명에 정확히 [{target_date_str}]이 포함된 PPT만 엄격하게 수집합니다...")
        print("-" * 60)

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)

        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)  # 최신순 정렬

        for message in messages:
            if message.Class != 43: continue

            try:
                # 메일 도착 시간이 오늘인지 확인
                msg_date = datetime.date(message.ReceivedTime.year, message.ReceivedTime.month,
                                         message.ReceivedTime.day)
                if msg_date < today: break
                if msg_date > today: continue

                sender_name = message.SenderName
                matched_expected_name = None

                # 우리 팀 담당자 메일인지 확인
                for expected_name in expected_teams.keys():
                    if expected_name in sender_name:
                        matched_expected_name = expected_name
                        break

                if not matched_expected_name: continue

                for att in message.Attachments:
                    filename = att.FileName

                    if filename.lower().endswith(('.ppt', '.pptx')):
                        # 🔒 [필터링 장치] 파일명에 오늘 날짜(예: 260521)가 무조건 있어야만 통과!
                        if target_date_str in filename:
                            time_str = message.ReceivedTime.strftime("%H%M%S")
                            safe_filename = f"{time_str}_{filename}"
                            save_path = os.path.join(save_folder, safe_filename)

                            att.SaveAsFile(save_path)
                            merge_order = expected_teams[matched_expected_name][0]
                            received_senders.add(matched_expected_name)
                            downloaded_files.append((merge_order, save_path, sender_name))

                            print(f"  ✅ [수집 성공] {filename} (보낸사람: {sender_name})")
                        else:
                            # 260514 등 과거 날짜를 보낸 경우 이곳으로 빠짐 (저장 안 됨)
                            wrong_date_senders.add(matched_expected_name)
                            print(f"  🚫 [수집 거부 - 날짜 틀림] {filename} (보낸사람: {sender_name})")

            except Exception:
                pass

        wrong_date_senders = wrong_date_senders - received_senders

        # ==========================================
        # [1.5단계] 미제출 및 실수한 팀 출력
        # ==========================================
        print("\n" + "=" * 60)
        missing_names = set(expected_teams.keys()) - received_senders

        if missing_names:
            print("🚨 [자료 확인 필요 부서 목록]")
            missing_list = [(expected_teams[name][0], expected_teams[name][1], name) for name in missing_names]
            missing_list.sort(key=lambda x: x[0])

            for order, team, name in missing_list:
                if name in wrong_date_senders:
                    # 파일명에 날짜를 잘못 쓴 부서에게 경고
                    print(f"   ❌ {team} ({name}) : 파일명 오류 (파일명에 '{target_date_str}'이 없습니다. 지난 자료인지 확인 요망)")
                else:
                    print(f"   ⚠️ {team} ({name}) : 미제출")
        else:
            print(f"🎉 모든 팀이 완벽하게 '{target_date_str}' 형식의 자료를 제출했습니다!")
        print("=" * 60)

        if not downloaded_files:
            print(f"\n병합할 {target_date_str} 기준 PPT 파일이 없습니다. 프로그램을 종료합니다.")
            return

        # ==========================================
        # [2단계] 지정된 순서에 맞춰 파일 병합
        # ==========================================
        print(f"\n2. 선별된 총 {len(downloaded_files)}개의 PPT 파일을 순서대로 병합합니다...")

        downloaded_files.sort(key=lambda x: x[0])

        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True

        base_presentation = ppt_app.Presentations.Open(downloaded_files[0][1])

        for file_info in downloaded_files[1:]:
            order, file_path, sender = file_info
            insert_index = base_presentation.Slides.Count
            base_presentation.Slides.InsertFromFile(file_path, insert_index)

        merged_filename = os.path.join(save_folder, "주간보고병합.pptx")

        if os.path.exists(merged_filename):
            try:
                os.remove(merged_filename)
            except Exception:
                pass

        base_presentation.SaveAs(merged_filename)
        base_presentation.Close()
        ppt_app.Quit()

        print(f"\n✨ '주간보고병합.pptx' 생성이 완료되었습니다!")
        print(f"📂 파일 위치: {merged_filename}")

    except Exception as e:
        print(f"실행 중 오류가 발생했습니다: {e}")


if __name__ == "__main__":
    download_and_merge_ppts()
