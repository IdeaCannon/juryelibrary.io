import tkinter as tk
from tkinterdnd2 import TkinterDnD
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os

tkdnd_path = os.path.abspath("tkdnd2.9.5")  # 또는 절대 경로 지정
os.environ['TKDND_LIBRARY'] = tkdnd_path

def generate_js(file_path):
    try:
        # 엑셀 읽기 및 필요한 열 추출
        df = pd.read_excel(file_path)
        required_columns = ['별치기호', '시작번호', '끝번호', '서가번호']
        df = df[required_columns]
        df['별치기호'] = df['별치기호'].fillna('')

        # 헤더 버전
        today_str = datetime.today().strftime('%Y-%m-%d-%H:%M')
        header_version_line = f"// @version      {today_str}"

        # location.js 앞부분
        rest_of_code1 = f"""
// ==UserScript==
// @name         주례열린 서가위치 출력
// @namespace    juryeopenlibrary
{header_version_line}"""
        rest_of_code2 = """\n// @description  도서관 서가위치&지도 표시
// @author       Sung Jong-Min
// @match        https://library.busan.go.kr/juryebooks/book/search/bookPrint
// @icon         https://www.google.com/s2/favicons?sz=64&domain=library.busan.go.kr
// @grant        none
// @run-at       document-start
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @updateURL    https://raw.githubusercontent.com/juryeopenlib/juryelibrary.io/refs/heads/main/location.js
// @downloadURL  https://raw.githubusercontent.com/juryeopenlib/juryelibrary.io/refs/heads/main/location.js
// ==/UserScript==

(function() {
    'use strict';
    function onDocumentReady(callback) {
        if (document.readyState === 'complete' || document.readyState === 'interactive') {
            callback();
        } else {
            document.addEventListener('DOMContentLoaded', callback);
        }
    }

    onDocumentReady(() => {
        // 프린트 함수 가로채기
        window.newPrint = window.print;
        window.print = () => {};

        // 윈도우 닫기 함수 가로채기
        window.newClose = window.close;
        window.close = () => {};

        $('colgroup col.col1').css('width', '50px');

        let trList = $('tr');
        console.log(trList);

        let 서명 = trList.eq(1).find('td:eq(1)').text();
        let 청구기호 = trList.eq(2).find('td:eq(1)').text();
        let 등록번호 = trList.eq(3).find('td:eq(1)').text();
        let 저자 = trList.eq(4).find('td:eq(1)').text();
        let 자료실 = trList.eq(5).find('td:eq(1)').text();

        // 기존 tr 삭제
        trList.eq(1).remove();
        trList.eq(2).remove();
        trList.eq(3).remove();
        trList.eq(4).remove();
        trList.eq(5).remove();
    """
        # 배열 변환
        indent = '        '  # 8칸 들여쓰기 기준
        array_lines = [f'{indent}const 서가번호표 = [']

        for i, row in enumerate(df.itertuples(index=False)):
            line = indent + '    { ' + ', '.join(
                f"{col}: '{getattr(row, col)}'" for col in required_columns
            ) + ' }'
            if i < len(df) - 1:
                line += ','  # 마지막 줄이 아니면 쉼표 추가
            array_lines.append(line)

        array_lines.append(f'{indent}];\n')
        array_lines_text = '\n'.join(array_lines)

        rest_of_code3 = """
        // 복합 번호 비교 함수
        function compareCompoundNumbers(num1, num2) {
            let parts1 = num1.split('-');
            let parts2 = num2.split('-');

            // 소수점 부분을 포함하여 각 파트를 비교
            for (let i = 0; i < Math.max(parts1.length, parts2.length); i++) {
                let part1 = parts1[i] ? parseFloat(parts1[i]) : 0; // 없으면 0으로 채움
                let part2 = parts2[i] ? parseFloat(parts2[i]) : 0;

                if (part1 < part2) return -1;
                if (part1 > part2) return 1;
            }
            return 0; // 완전히 같으면 0
        }

        // 청구기호에서 별치기호와 번호 추출
        let 청구기호부분들 = 청구기호.split(' ');

        // 청구기호가 별치기호 없이 숫자만 있을 경우를 처리
        let 별치기호 = 청구기호부분들.length === 2 ? 청구기호부분들[0] : ''; // 별치기호가 없으면 빈칸
        let 번호 = 청구기호부분들.length === 2 ? 청구기호부분들[1] : 청구기호부분들[0]; // 번호는 항상 두 번째 요소

        // 자료실 값에 따른 서가위치 설정
        let 서가위치 = '서가번호를 찾을 수 없습니다';
        if (자료실.includes('신간도서(지하1층)')) {
            서가위치 = '지하1층-신간도서 코너';
        } else if (자료실.includes('신간도서(2층)')) {
            서가위치 = '2층-신간도서 코너';
        } else if (자료실.includes('신간도서(3층)')) {
            서가위치 = '3층-신간도서 코너';
        } else if (자료실.includes('북큐레이션(지하1층)')) {
            서가위치 = '지하1층-EV앞';
        } else if (자료실.includes('북큐레이션(2층)')) {
            서가위치 = '2층-신간코너 옆';
        } else if (자료실.includes('북큐레이션(3층EV앞)')) {
            서가위치 = '3층-EV앞';
        } else if (자료실.includes('북큐레이션(3층정문)')) {
            서가위치 = '3층-입구 옆면';
        } else {
            // 서가번호 찾기
            for (let i = 0; i < 서가번호표.length; i++) {
                let 항목 = 서가번호표[i];
                if (항목.별치기호 === 별치기호 && compareCompoundNumbers(번호, 항목.시작번호) >= 0 && compareCompoundNumbers(번호, 항목.끝번호) <= 0) {
                    서가위치 = 항목.서가번호;
                    break;
                }
            }
        }


        // 표 정렬을 위한 새로운 tr 추가
        trList.eq(0).find('td:eq(0)').html("<div style='border-top: 1px dashed black;'></div>");

        let title = `<tr style='font-size: 14px; text-align: justify; font-weight: bold; font-family: 돋움' class='first td1'>
                        <td style='font-size: 14px; font-weight: bold; font-family: 돋움'>제&nbsp;&nbsp;&nbsp;&nbsp;목:</td>
                        <td style='font-size: 14px; font-weight: bold; font-family: 돋움' class='last td2'>${서명}</td>
                     </tr>`;
        $('tr').eq(1).before(title);

        let CallNumber = `<tr style='font-size: 14px; text-align: justify; font-weight: bold; font-family: 돋움' class='first td1'>
                        <td style='font-size: 14px; font-weight: bold; font-family: 돋움'>청구기호:</td>
                        <td style='font-size: 14px; font-weight: bold; font-family: 돋움' class='last td2'>${청구기호}</td>
                     </tr>`;
        $('tr').eq(2).before(CallNumber);

        let RegiNumber = `<tr style='font-size: 14px; text-align: justify; font-weight: bold; font-family: 돋움' class='first td1'>
                        <td style='font-size: 14px; font-weight: bold; font-family: 돋움'>등록번호:</td>
                        <td style='font-size: 14px; font-weight: bold; font-family: 돋움' class='last td2'>${등록번호}</td>
                     </tr>`;
        $('tr').eq(3).before(RegiNumber);

        let Author = `<tr style='font-size: 14px; text-align: justify; font-weight: bold; font-family: 돋움' class='first td1'>
                        <td style='font-size: 14px; font-weight: bold; font-family: 돋움'>저&nbsp;&nbsp;&nbsp;&nbsp;자:</td>
                        <td style='font-size: 14px; font-weight: bold; font-family: 돋움' class='last td2'>${저자}</td>
                     </tr>`;
        $('tr').eq(4).before(Author);

        자료실 = 자료실.replace('[주례열린]', '').replace('종합자료실(지하1층)', '종합자료실').replace('어린이자료실1(1층)', '어린이자료실1').replace('어린이자료실2(2층)', '어린이자료실2').replace('유아자료실(2층)', '유아자료실').replace('열린자료실(3층)', '열린자료실').trim();
        let room = `<tr style='font-size: 14px; text-align: justify; font-weight: bold; font-family: 돋움' class='first td1'>
                        <td style='font-size: 14px; font-weight: bold; font-family: 돋움'>자&nbsp;료&nbsp;실: </td>
                        <td style='font-size: 14px; font-weight: bold; font-family: 돋움' class='last td2'>${자료실}</td>
                     </tr>`;
        $('tr').eq(5).before(room);

        trList.eq(6).find('td:eq(0)').html("")


        let lastTr = $('tr:last');
        lastTr.before(`<tr style='font-size: 14px; text-align: justify; font-weight: bold;font-family: 돋움' class='first td1'><td style='font-size: 14px; font-weight: bold;font-family: 돋움 '>서가위치: </td><td style="font-size: 14px; font-weight: bold;font-family: 돋움 " class="last td2" >${서가위치}</td></tr>`);

        // 서가위치에 따라 이미지 호출
        let imageUrl;

        if (서가위치 === '어린이자료실1(1층)-01') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/1%EC%B8%B5-1.png';
        } else if (서가위치 === '어린이자료실1(1층)-02') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/1%EC%B8%B5-2.png';
        } else if (서가위치 === '어린이자료실1(1층)-03') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/1%EC%B8%B5-3.png';
        } else if (서가위치 === '어린이자료실1(1층)-04') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/1%EC%B8%B5-4.png';
        } else if (서가위치 === '어린이자료실1(1층)-05') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/1%EC%B8%B5-5.png';
        } else if (서가위치 === '어린이자료실1(1층)-M') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/1%EC%B8%B5-M.png';
        } else if (서가위치 === '어린이자료실2(2층)-06') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-6.png';
        } else if (서가위치 === '어린이자료실2(2층)-07') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-7.png';
        } else if (서가위치 === '어린이자료실2(2층)-08') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-8.png';
        } else if (서가위치 === '어린이자료실2(2층)-09') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-9.png';
        } else if (서가위치 === '어린이자료실2(2층)-10') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-10.png';
        } else if (서가위치 === '어린이자료실2(2층)-11') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-11.png';
        } else if (서가위치 === '어린이자료실2(2층)-12') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-12.png';
        } else if (서가위치 === '어린이자료실2(2층)-13') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-13.png';
        } else if (서가위치 === '어린이자료실2(2층)-14') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-14.png';
        } else if (서가위치 === '어린이자료실2(2층)-15') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-15.png';
        } else if (서가위치 === '어린이자료실2(2층)-16') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-16.png';
        } else if (서가위치 === '어린이자료실2(2층)-17') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-17.png';
        } else if (서가위치 === '어린이자료실2(2층)-18') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-18.png';
        } else if (서가위치 === '어린이자료실2(2층)-19') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-19.png';
        } else if (서가위치 === '어린이자료실2(2층)-20') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-20.png';
        } else if (서가위치 === '어린이자료실2(2층)-21') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-21.png';
        } else if (서가위치 === '어린이자료실2(2층)-22') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-22.png';
        } else if (서가위치 === '어린이자료실2(2층)-23') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-23.png';
        } else if (서가위치 === '어린이자료실2(2층)-D') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-D.png';
        } else if (서가위치 === '어린이자료실2(2층)-H') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-H.png';
        } else if (서가위치 === '어린이자료실2(2층)-I') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-I.png';
        } else if (서가위치 === '어린이자료실2(2층)-J') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-J.png';
        } else if (서가위치 === '어린이자료실2(2층)-K') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-K.png';
        } else if (서가위치 === '어린이자료실2(2층)-L') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-L.png';
        } else if (서가위치 === '열린자료실(3층)-A') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/3%EC%B8%B5-A.png';
        } else if (서가위치 === '열린자료실(3층)-B') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/3%EC%B8%B5-B.png';
        } else if (서가위치 === '열린자료실(3층)-C') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/3%EC%B8%B5-C.png';
        } else if (서가위치 === '열린자료실(3층)-청춘쉼터') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/3%EC%B8%B5-%ED%8A%B9%ED%99%94.png';
        } else if (서가위치 === '열린자료실(3층)-청춘진로') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/3%EC%B8%B5-%ED%8A%B9%ED%99%94.png';
        } else if (서가위치 === '유아자료실(2층)-E') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-E.png';
        } else if (서가위치 === '유아자료실(2층)-F') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-F.png';
        } else if (서가위치 === '유아자료실(2층)-G') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-G.png';
        } else if (서가위치 === '종합자료실(지하1층)-01') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-1.png';
        } else if (서가위치 === '종합자료실(지하1층)-02') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-2.png';
        } else if (서가위치 === '종합자료실(지하1층)-03') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-3.png';
        } else if (서가위치 === '종합자료실(지하1층)-04') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-4.png';
        } else if (서가위치 === '종합자료실(지하1층)-05') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-5.png';
        } else if (서가위치 === '종합자료실(지하1층)-06') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-6.png';
        } else if (서가위치 === '종합자료실(지하1층)-07') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-7.png';
        } else if (서가위치 === '종합자료실(지하1층)-08') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-8.png';
        } else if (서가위치 === '종합자료실(지하1층)-09') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-9.png';
        } else if (서가위치 === '종합자료실(지하1층)-10') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-10.png';
        } else if (서가위치 === '종합자료실(지하1층)-11') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-11.png';
        } else if (서가위치 === '종합자료실(지하1층)-12') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-12.png';
        } else if (서가위치 === '종합자료실(지하1층)-13') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-13.png';
        } else if (서가위치 === '종합자료실(지하1층)-14') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-14.png';
        } else if (서가위치 === '종합자료실(지하1층)-15') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-15.png';
        } else if (서가위치 === '종합자료실(지하1층)-16') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-16.png';
        } else if (서가위치 === '종합자료실(지하1층)-17') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-17.png';
        } else if (서가위치 === '종합자료실(지하1층)-18') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-18.png';
        } else if (서가위치 === '종합자료실(지하1층)-19') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-19.png';
        } else if (서가위치 === '종합자료실(지하1층)-20') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-20.png';
        } else if (서가위치 === '종합자료실(지하1층)-21') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-21.png';
        } else if (서가위치 === '종합자료실(지하1층)-N') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-N.png';
        } else if (서가위치 === '종합자료실(지하1층)-O') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-O.png';
        } else if (서가위치 === '종합자료실(지하1층)-P') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-P.png';
        } else if (서가위치 === '종합자료실(지하1층)-Q') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-Q.png';
        } else if (서가위치 === '종합자료실(지하1층)-R') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-R.png';
        } else if (서가위치 === '종합자료실(지하1층)-S') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-S.png';
        } else if (서가위치 === '종합자료실(지하1층)-T') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%981%EC%B8%B5-T.png';
        } else if (서가위치 === '지하1층-신간도서 코너') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%A7%80%ED%95%98%201%EC%B8%B5-%EC%8B%A0%EA%B0%84.png';
        } else if (서가위치 === '2층-신간도서 코너') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/2%EC%B8%B5-%EC%8B%A0%EA%B0%84.png';
        } else if (서가위치 === '3층-신간도서 코너') {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/3%EC%B8%B5-%EC%8B%A0%EA%B0%84.png';
        } else {
            imageUrl = 'https://juryeopenlib.github.io/juryelibrary.io/maps/%EC%9D%B4%EB%AF%B8%EC%A7%80%20%EC%B6%94%EA%B0%80%EC%A4%91.png';//  기본이미지
        }

        // 이미지 추가
        lastTr.before(`
            <tr>
                <td colspan="2" style="text-align: center;">
                    <img src="${imageUrl}"
                         style="width: 250px; float: left;"/>
                </td>
            </tr>
        `);


        // 이미지 로드가 완료되면 창 닫기
        let imageCount = $('img').length;
        let loadedImageCount = 0;

        $('img').on('load', function() {
            loadedImageCount++;
            if (loadedImageCount === imageCount) {
                window.newPrint();
                setTimeout(window.newClose, 1000); // 1초 지연 후 창 닫기
            }
        });
    });
})();

"""

        # JS 파일 합치기
        full_js_code = rest_of_code1 + rest_of_code2 + '\n' + array_lines_text + rest_of_code3

        # 저장
        output_path = os.path.join(os.path.dirname(file_path), 'location.js')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_js_code)

        messagebox.showinfo("완료", f"location.js 파일이 생성되었습니다:\n{output_path}")

    except Exception as e:
        messagebox.showerror("오류", f"에러 발생: {str(e)}")

# 드래그 앤 드롭 처리
def handle_drop(event):
    dropped_file = event.data.strip('{}')
    if dropped_file.lower().endswith('.xlsx'):
        generate_js(dropped_file)
    else:
        messagebox.showerror("오류", "엑셀 파일만 가능합니다 (.xlsx)")

# GUI 구성
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    root = TkinterDnD.Tk()
    root.title("서가번호 location.js 생성기")
    root.geometry("300x100")

    label = tk.Label(root, text="📂 서가 배열 엑셀 파일을 \n여기에 끌여당겨주세요\n드래그 앤 드롭", font=("돋움", 14))
    label.pack(expand=True)
    label.drop_target_register(DND_FILES)
    label.dnd_bind('<<Drop>>', handle_drop)
    root.mainloop()
except ImportError:
    tk.messagebox.showerror("모듈 오류", "drag & drop 기능을 위해 tkinterdnd2 모듈이 필요합니다.\n\n📦 pip install tkinterdnd2")


root.mainloop()