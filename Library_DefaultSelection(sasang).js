// ==UserScript==
// @name         도서관 기본 선택(영도도서관)
// @namespace    https://library.busan.go.kr/
// @version      1.01
// @description  도서관 자료 검색 시 도서관에 맞게 기본선택해줍니다.
// @author       hideD
// @match        https://library.busan.go.kr/*/book/search/collectionOfMaterials*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=library.busan.go.kr
// @grant        none
// @updateURL    https://raw.githubusercontent.com/IdeaCannon/juryelibrary.io/refs/heads/main/Library_DefaultSelection(sasang).js
// @downloadURL  https://raw.githubusercontent.com/IdeaCannon/juryelibrary.io/refs/heads/main/Library_DefaultSelection(sasang).js
// ==/UserScript==

(function() {
    'use strict';

    const library_map = {
        ssbooks: 'AM', //사상도서관
    };

    const option_key = '#search_managecode option'; //선택 가능한 옵션들
    const select_key = '#search_managecode';

    //document.location.pathname == '/gsbooks/book/search/collectionOfMaterials'
    const path_list = document.location.pathname.split('/');
    //['', 'gsbooks', 'book', 'search', 'collectionOfMaterials']
    if(path_list.length < 2){
        console.error('도서관 주소가 올바르지 않습니다.');
        return;
    }


    const library_key = path_list[1];
    if(!(library_key in library_map)){
        console.error(`도서관 주소 ${library_key}가 library_map에 없습니다`);
        return;
    }
    const library_value = library_map[library_key]; //대문자값 추출

    const valid_values = new Map();
    for(const item of document.querySelectorAll(option_key)){
        valid_values.set(item.value, item.innerText);
    }

    if(valid_values == 0){
        console.error(`${option_key}에 선택 가능한 값이 없습니다. query를 확인해주세요`);
        return;
    }

    //실제로 선택 가능한 값인가를 확인하기 위해
    const library_hangul_name = valid_values.get(library_value);

    if(!library_hangul_name){
        console.error(`${[...valid_values.values()].join(', ')} 중에 ${library_value}가 없습니다. library_map을 확인해주세요.`);
        return;
    }

    const select_target = document.querySelector(select_key);
    if(!select_target){
        console.error(`${select_key}가 잘못되었습니다. 다른 값을 찾아주세요`);
        return;
    }
    select_target.value = library_value;
    console.log(`${library_hangul_name}을 기본 선택합니다`);
})();
