// ==UserScript==
// @name         주례열린 서가위치 출력
// @namespace    juryeopenlibrary
// @version      2024-12-10
// @description  도서관 서가위치&지도 표시
// @author       IdeaCannon(Assistant by HideD)
// @match        https://library.busan.go.kr/juryebooks/book/search/bookPrint
// @icon         https://www.google.com/s2/favicons?sz=64&domain=library.busan.go.kr
// @grant        none
// @run-at       document-start
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @updateURL    https://raw.githubusercontent.com/IdeaCannon/juryelibrary.io/refs/heads/main/location.js
// @downloadURL  https://raw.githubusercontent.com/IdeaCannon/juryelibrary.io/refs/heads/main/location.js
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

        const 서가번호표 = [
            { 별치기호: '더책', 시작번호: '069-1', 끝번호: '999.999', 서가번호: '유아자료실(2층)-E' },
            { 별치기호: '아동', 시작번호: '001-1', 끝번호: '29.999', 서가번호: '어린이자료실2(2층)-06' },
            { 별치기호: '아동', 시작번호: '030-1', 끝번호: '031-13-999', 서가번호: '어린이자료실2(2층)-07' },
            { 별치기호: '아동', 시작번호: '031-14-1', 끝번호: '082-9-999', 서가번호: '어린이자료실2(2층)-08' },
            { 별치기호: '아동', 시작번호: '082-10-1', 끝번호: '99.999', 서가번호: '어린이자료실2(2층)-09' },
            { 별치기호: '아동', 시작번호: '100-1', 끝번호: '199.999', 서가번호: '어린이자료실2(2층)-L' },
            { 별치기호: '아동', 시작번호: '300-1-1', 끝번호: '319.999', 서가번호: '어린이자료실2(2층)-K' },
            { 별치기호: '아동', 시작번호: '320-1', 끝번호: '341.999', 서가번호: '어린이자료실2(2층)-J' },
            { 별치기호: '아동', 시작번호: '342.1-1', 끝번호: '399.999', 서가번호: '어린이자료실2(2층)-I' },
            { 별치기호: '아동', 시작번호: '400-1-1', 끝번호: '408-3-50', 서가번호: '어린이자료실2(2층)-16' },
            { 별치기호: '아동', 시작번호: '408-3-51', 끝번호: '410-10', 서가번호: '어린이자료실2(2층)-17' },
            { 별치기호: '아동', 시작번호: '410-10-1', 끝번호: '451.31-999', 서가번호: '어린이자료실2(2층)-18' },
            { 별치기호: '아동', 시작번호: '451.32-1', 끝번호: '499.999', 서가번호: '어린이자료실2(2층)-19' },
            { 별치기호: '아동', 시작번호: '500-1', 끝번호: '599.999', 서가번호: '어린이자료실2(2층)-20' },
            { 별치기호: '아동', 시작번호: '600-1', 끝번호: '799.999', 서가번호: '어린이자료실2(2층)-H' },
            { 별치기호: '아동', 시작번호: '800-1', 끝번호: '808.3-9-200', 서가번호: '어린이자료실2(2층)-D' },
            { 별치기호: '아동', 시작번호: '808.3-9-201', 끝번호: '811.8-59', 서가번호: '어린이자료실2(2층)-10' },
            { 별치기호: '아동', 시작번호: '811.8-60', 끝번호: '813.8-969', 서가번호: '어린이자료실2(2층)-11' },
            { 별치기호: '아동', 시작번호: '813.8-2561', 끝번호: '833.8-183-999', 서가번호: '어린이자료실2(2층)-13' },
            { 별치기호: '아동', 시작번호: '813.8-970', 끝번호: '813.8-2560', 서가번호: '어린이자료실2(2층)-12' },
            { 별치기호: '아동', 시작번호: '833.8-184', 끝번호: '843-380', 서가번호: '어린이자료실2(2층)-14' },
            { 별치기호: '아동', 시작번호: '843-381', 끝번호: '899.999', 서가번호: '어린이자료실2(2층)-15' },
            { 별치기호: '아동', 시작번호: '900-1', 끝번호: '911-30-999', 서가번호: '어린이자료실2(2층)-21' },
            { 별치기호: '아동', 시작번호: '911-31-1', 끝번호: '989.999', 서가번호: '어린이자료실2(2층)-22' },
            { 별치기호: '아동', 시작번호: '990-1', 끝번호: '999.999', 서가번호: '어린이자료실2(2층)-23' },
            { 별치기호: '아동영어', 시작번호: '0', 끝번호: '747-3-1', 서가번호: '어린이자료실1(1층)-02' },
            { 별치기호: '아동영어', 시작번호: '747-3-1-1', 끝번호: '843-37', 서가번호: '어린이자료실1(1층)-03' },
            { 별치기호: '아동영어', 시작번호: '843-183', 끝번호: '999.999', 서가번호: '어린이자료실1(1층)-05' },
            { 별치기호: '아동영어', 시작번호: '843-38', 끝번호: '843-182-999', 서가번호: '어린이자료실1(1층)-04' },
            { 별치기호: '영어', 시작번호: '0', 끝번호: '999.999', 서가번호: '어린이자료실1(1층)-M' },
            { 별치기호: '유아', 시작번호: '000-1', 끝번호: '799.999', 서가번호: '유아자료실(2층)-F' },
            { 별치기호: '유아', 시작번호: '900-1', 끝번호: '999.999', 서가번호: '유아자료실(2층)-F' },
            { 별치기호: '유아', 시작번호: '808-1-1', 끝번호: '899.999', 서가번호: '유아자료실(2층)-G' },
            { 별치기호: '팝업', 시작번호: '001-1', 끝번호: '999.999', 서가번호: '유아자료실(2층)-G' },
            { 별치기호: '일본', 시작번호: '0', 끝번호: '999.999', 서가번호: '어린이자료실1(1층)-02' },
            { 별치기호: '중국', 시작번호: '0', 끝번호: '999.999', 서가번호: '어린이자료실1(1층)-01' },
            { 별치기호: '청춘쉼터', 시작번호: '0', 끝번호: '999.999', 서가번호: '열린자료실(3층)-청춘쉼터' },
            { 별치기호: '청춘진로', 시작번호: '0', 끝번호: '999.999', 서가번호: '열린자료실(3층)-청춘진로' },
            { 별치기호: '캄보디아', 시작번호: '0', 끝번호: '999.999', 서가번호: '어린이자료실1(1층)-01' },
            { 별치기호: '베트남', 시작번호: '0', 끝번호: '999.999', 서가번호: '어린이자료실1(1층)-01' },
            { 별치기호: '태국', 시작번호: '0', 끝번호: '999.999', 서가번호: '어린이자료실1(1층)-01' },
            { 별치기호: '큰글자', 시작번호: '0', 끝번호: '999.999', 서가번호: '열린자료실(3층)-A' },
            { 별치기호: '아트', 시작번호: '0', 끝번호: '999.999', 서가번호: '2층 계단 벽면서가' },
            { 별치기호: '쉬운책', 시작번호: '0', 끝번호: '999.999', 서가번호: '열린자료실(3층)-청춘쉼터' },
            { 별치기호: '점자', 시작번호: '0', 끝번호: '999.999', 서가번호: '열린자료실(3층)-청춘쉼터' },
            { 별치기호: '참고도서', 시작번호: '0', 끝번호: '999.999', 서가번호: '지하 1층 데스크 문의' },
            { 별치기호: '', 시작번호: '001-1', 끝번호: '26.999', 서가번호: '종합자료실(지하1층)-P' },
            { 별치기호: '', 시작번호: '027-1', 끝번호: '99.999', 서가번호: '종합자료실(지하1층)-Q' },
            { 별치기호: '', 시작번호: '100-1', 끝번호: '181.71-59', 서가번호: '종합자료실(지하1층)-R' },
            { 별치기호: '', 시작번호: '181.72', 끝번호: '199.999', 서가번호: '종합자료실(지하1층)-S' },
            { 별치기호: '', 시작번호: '200-1', 끝번호: '219.999', 서가번호: '종합자료실(지하1층)-T' },
            { 별치기호: '', 시작번호: '220-1', 끝번호: '299.999', 서가번호: '종합자료실(지하1층)-N' },
            { 별치기호: '', 시작번호: '300-1', 끝번호: '321.341-999', 서가번호: '종합자료실(지하1층)-01' },
            { 별치기호: '', 시작번호: '321.342-1', 끝번호: '325.1-52', 서가번호: '종합자료실(지하1층)-02' },
            { 별치기호: '', 시작번호: '325.1-53', 끝번호: '325.211-359', 서가번호: '종합자료실(지하1층)-03' },
            { 별치기호: '', 시작번호: '325.211-360', 끝번호: '325.571-8', 서가번호: '종합자료실(지하1층)-04' },
            { 별치기호: '', 시작번호: '325.571-9', 끝번호: '327.85-999', 서가번호: '종합자료실(지하1층)-05' },
            { 별치기호: '', 시작번호: '327.851-1-1', 끝번호: '327.87-127', 서가번호: '종합자료실(지하1층)-06' },
            { 별치기호: '', 시작번호: '327.87-128', 끝번호: '331.5412-34', 서가번호: '종합자료실(지하1층)-07' },
            { 별치기호: '', 시작번호: '331.5412-35', 끝번호: '339.999', 서가번호: '종합자료실(지하1층)-08' },
            { 별치기호: '', 시작번호: '340-1', 끝번호: '359.999', 서가번호: '종합자료실(지하1층)-09' },
            { 별치기호: '', 시작번호: '360-1', 끝번호: '399.999', 서가번호: '종합자료실(지하1층)-O' },
            { 별치기호: '', 시작번호: '400-1', 끝번호: '472.52-4', 서가번호: '종합자료실(지하1층)-10' },
            { 별치기호: '', 시작번호: '472.52-5', 끝번호: '513.9708-999', 서가번호: '종합자료실(지하1층)-11' },
            { 별치기호: '', 시작번호: '513.9709-1', 끝번호: '592.27-34', 서가번호: '종합자료실(지하1층)-12' },
            { 별치기호: '', 시작번호: '592.27-35', 끝번호: '599.999', 서가번호: '종합자료실(지하1층)-13' },
            { 별치기호: '', 시작번호: '600-1', 끝번호: '669.999', 서가번호: '열린자료실(3층)-C' },
            { 별치기호: '', 시작번호: '670-1', 끝번호: '699.999', 서가번호: '열린자료실(3층)-B' },
            { 별치기호: '', 시작번호: '700-1', 끝번호: '799.999', 서가번호: '종합자료실(지하1층)-14' },
            { 별치기호: '', 시작번호: '800-1', 끝번호: '808-14-98', 서가번호: '종합자료실(지하1층)-16' },
            { 별치기호: '', 시작번호: '808-14-99', 끝번호: '813.7-519', 서가번호: '종합자료실(지하1층)-17' },
            { 별치기호: '', 시작번호: '813.7-520', 끝번호: '817.999', 서가번호: '종합자료실(지하1층)-18' },
            { 별치기호: '', 시작번호: '818-1', 끝번호: '829.999', 서가번호: '종합자료실(지하1층)-19' },
            { 별치기호: '', 시작번호: '830-1', 끝번호: '843-389', 서가번호: '종합자료실(지하1층)-20' },
            { 별치기호: '', 시작번호: '843-390', 끝번호: '899.999', 서가번호: '종합자료실(지하1층)-21' },
            { 별치기호: '', 시작번호: '900-1', 끝번호: '909.999', 서가번호: '종합자료실(지하1층)-14' },
            { 별치기호: '', 시작번호: '910-1', 끝번호: '919.999', 서가번호: '종합자료실(지하1층)-15' },
            { 별치기호: '', 시작번호: '920-1', 끝번호: '999.999', 서가번호: '종합자료실(지하1층)-16' }

        ];

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
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/1%EC%B8%B5-1.png';
        } else if (서가위치 === '어린이자료실1(1층)-02') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/1%EC%B8%B5-2.png';
        } else if (서가위치 === '어린이자료실1(1층)-03') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/1%EC%B8%B5-3.png';
        } else if (서가위치 === '어린이자료실1(1층)-04') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/1%EC%B8%B5-4.png';
        } else if (서가위치 === '어린이자료실1(1층)-05') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/1%EC%B8%B5-5.png';
        } else if (서가위치 === '어린이자료실1(1층)-M') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/1%EC%B8%B5-M.png';
        } else if (서가위치 === '어린이자료실2(2층)-06') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-6.png';
        } else if (서가위치 === '어린이자료실2(2층)-07') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-7.png';
        } else if (서가위치 === '어린이자료실2(2층)-08') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-8.png';
        } else if (서가위치 === '어린이자료실2(2층)-09') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-9.png';
        } else if (서가위치 === '어린이자료실2(2층)-10') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-10.png';
        } else if (서가위치 === '어린이자료실2(2층)-11') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-11.png';
        } else if (서가위치 === '어린이자료실2(2층)-12') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-12.png';
        } else if (서가위치 === '어린이자료실2(2층)-13') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-13.png';
        } else if (서가위치 === '어린이자료실2(2층)-14') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-14.png';
        } else if (서가위치 === '어린이자료실2(2층)-15') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-15.png';
        } else if (서가위치 === '어린이자료실2(2층)-16') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-16.png';
        } else if (서가위치 === '어린이자료실2(2층)-17') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-17.png';
        } else if (서가위치 === '어린이자료실2(2층)-18') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-18.png';
        } else if (서가위치 === '어린이자료실2(2층)-19') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-19.png';
        } else if (서가위치 === '어린이자료실2(2층)-20') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-20.png';
        } else if (서가위치 === '어린이자료실2(2층)-21') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-21.png';
        } else if (서가위치 === '어린이자료실2(2층)-22') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-22.png';
        } else if (서가위치 === '어린이자료실2(2층)-23') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-23.png';
        } else if (서가위치 === '어린이자료실2(2층)-D') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-D.png';
        } else if (서가위치 === '어린이자료실2(2층)-H') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-H.png';
        } else if (서가위치 === '어린이자료실2(2층)-I') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-I.png';
        } else if (서가위치 === '어린이자료실2(2층)-J') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-J.png';
        } else if (서가위치 === '어린이자료실2(2층)-K') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-K.png';
        } else if (서가위치 === '어린이자료실2(2층)-L') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-L.png';
        } else if (서가위치 === '열린자료실(3층)-A') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/3%EC%B8%B5-A.png';
        } else if (서가위치 === '열린자료실(3층)-B') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/3%EC%B8%B5-B.png';
        } else if (서가위치 === '열린자료실(3층)-C') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/3%EC%B8%B5-C.png';
        } else if (서가위치 === '열린자료실(3층)-청춘쉼터') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/3%EC%B8%B5-%ED%8A%B9%ED%99%94.png';
        } else if (서가위치 === '열린자료실(3층)-청춘진로') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/3%EC%B8%B5-%ED%8A%B9%ED%99%94.png';
        } else if (서가위치 === '유아자료실(2층)-E') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-E.png';
        } else if (서가위치 === '유아자료실(2층)-F') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-F.png';
        } else if (서가위치 === '유아자료실(2층)-G') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-G.png';
        } else if (서가위치 === '종합자료실(지하1층)-01') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-1.png';
        } else if (서가위치 === '종합자료실(지하1층)-02') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-2.png';
        } else if (서가위치 === '종합자료실(지하1층)-03') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-3.png';
        } else if (서가위치 === '종합자료실(지하1층)-04') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-4.png';
        } else if (서가위치 === '종합자료실(지하1층)-05') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-5.png';
        } else if (서가위치 === '종합자료실(지하1층)-06') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-6.png';
        } else if (서가위치 === '종합자료실(지하1층)-07') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-7.png';
        } else if (서가위치 === '종합자료실(지하1층)-08') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-8.png';
        } else if (서가위치 === '종합자료실(지하1층)-09') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-9.png';
        } else if (서가위치 === '종합자료실(지하1층)-10') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-10.png';
        } else if (서가위치 === '종합자료실(지하1층)-11') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-11.png';
        } else if (서가위치 === '종합자료실(지하1층)-12') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-12.png';
        } else if (서가위치 === '종합자료실(지하1층)-13') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-13.png';
        } else if (서가위치 === '종합자료실(지하1층)-14') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-14.png';
        } else if (서가위치 === '종합자료실(지하1층)-15') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-15.png';
        } else if (서가위치 === '종합자료실(지하1층)-16') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-16.png';
        } else if (서가위치 === '종합자료실(지하1층)-17') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-17.png';
        } else if (서가위치 === '종합자료실(지하1층)-18') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-18.png';
        } else if (서가위치 === '종합자료실(지하1층)-19') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-19.png';
        } else if (서가위치 === '종합자료실(지하1층)-20') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-20.png';
        } else if (서가위치 === '종합자료실(지하1층)-21') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-21.png';
        } else if (서가위치 === '종합자료실(지하1층)-N') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-N.png';
        } else if (서가위치 === '종합자료실(지하1층)-O') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-O.png';
        } else if (서가위치 === '종합자료실(지하1층)-P') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-P.png';
        } else if (서가위치 === '종합자료실(지하1층)-Q') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-Q.png';
        } else if (서가위치 === '종합자료실(지하1층)-R') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-R.png';
        } else if (서가위치 === '종합자료실(지하1층)-S') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-S.png';
        } else if (서가위치 === '종합자료실(지하1층)-T') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%981%EC%B8%B5-T.png';
        } else if (서가위치 === '지하1층-신간도서 코너') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%A7%80%ED%95%98%201%EC%B8%B5-%EC%8B%A0%EA%B0%84.png';
        } else if (서가위치 === '2층-신간도서 코너') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/2%EC%B8%B5-%EC%8B%A0%EA%B0%84.png';
        } else if (서가위치 === '3층-신간도서 코너') {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/3%EC%B8%B5-%EC%8B%A0%EA%B0%84.png';
        } else {
            imageUrl = 'https://ideacannon.github.io/juryelibrary.io/%EC%9D%B4%EB%AF%B8%EC%A7%80%20%EC%B6%94%EA%B0%80%EC%A4%91.png';//  기본이미지
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
