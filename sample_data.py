"""
기본 샘플 영수증(영수증.pdf)에 대한 정밀 분석 데이터
Gemini API 키가 아직 설정되지 않았거나 오프라인 시연 모드일 때 즉각적으로 동작할 수 있도록 지원합니다.
"""

SAMPLE_RECEIPTS = [
    {
        "page_index": 1,
        "receipt_sub_index": 1,
        "seq": 1,
        "date": "20261001",
        "time": "12:50:09",
        "store_name": "개포행복 파리바게뜨",
        "purpose": "이정복 보건소",  # 수기 메모 내용
        "items": [
            {"name": "호두파운드1호", "quantity": 2, "price": 31800},
            {"name": "말차)명가카스테라", "quantity": 1, "price": 15900}
        ],
        "total_quantity": 3,
        "attendees": "3명",
        "card_last4": "9716",
        "amount": 47700,
        "user_name": "",  # 영수증 실물에서 확인 불가 시 공란 처리
        "department": "",
        "notes": "정상 결제 / 삼성비자카드 / 수기메모: 이정복 보건소"
    },
    {
        "page_index": 1,
        "receipt_sub_index": 2,
        "seq": 2,
        "date": "20261001",
        "time": "12:50:33",
        "store_name": "개포행복 파리바게뜨",
        "purpose": "이정복 보건소",  # 수기 메모 내용
        "items": [
            {"name": "말차)명가카스테라", "quantity": 1, "price": 15900},
            {"name": "호두파운드1호", "quantity": 2, "price": 31800}
        ],
        "total_quantity": 3,
        "attendees": "3명",
        "card_last4": "9716",
        "amount": 47700,
        "user_name": "",  # 영수증 실물에서 확인 불가 시 공란 처리
        "department": "",
        "notes": "정상 결제 / 삼성비자카드 / 수기메모: 이정복 보건소"
    },
    {
        "page_index": 2,
        "receipt_sub_index": 1,
        "seq": 3,
        "date": "20261006",
        "time": "18:04:52",
        "store_name": "개포행복 파리바게뜨",
        "purpose": "영등포 보건소",  # 수기 메모 내용
        "items": [
            {"name": "천혜향)명가카스테라", "quantity": 2, "price": 31800},
            {"name": "호두파운드1호", "quantity": 2, "price": 31800},
            {"name": "말차)명가카스테라", "quantity": 1, "price": 15900},
            {"name": "만월전병", "quantity": 1, "price": 19900},
            {"name": "세트)브라우니쿠키", "quantity": 2, "price": 11800}
        ],
        "total_quantity": 8,
        "attendees": "8명",
        "card_last4": "9716",
        "amount": 100100,  # KT할인 -11,100원 적용
        "user_name": "",  # 영수증 실물에서 확인 불가 시 공란 처리
        "department": "",
        "notes": "KT 멤버십 할인(-11,100원) 적용 / 수기메모: 영등포 보건소"
    },
    {
        "page_index": 2,
        "receipt_sub_index": 2,
        "seq": 4,
        "date": "20261006",
        "time": "18:07:27",
        "store_name": "개포행복 파리바게뜨",
        "purpose": "동작구 보건소",  # 수기 메모 내용
        "items": [
            {"name": "벌꿀)명가카스테라", "quantity": 1, "price": 15900},
            {"name": "호두파운드1호", "quantity": 1, "price": 15900},
            {"name": "선물류)복덕밤만주", "quantity": 1, "price": 19500}
        ],
        "total_quantity": 3,
        "attendees": "3명",
        "card_last4": "9716",
        "amount": 51300,
        "user_name": "",  # 영수증 실물에서 확인 불가 시 공란 처리
        "department": "",
        "notes": "정상 결제 / 수기메모: 동작구 보건소"
    },
    {
        "page_index": 3,
        "receipt_sub_index": 1,
        "seq": 5,
        "date": "20261006",
        "time": "18:08:59",
        "store_name": "개포행복 파리바게뜨",
        "purpose": "양주시 보건소",  # 수기 메모 내용
        "items": [
            {"name": "말차)명가카스테라", "quantity": 1, "price": 15900},
            {"name": "호두파운드1호", "quantity": 1, "price": 15900},
            {"name": "제주메밀카스테라(K파바)", "quantity": 1, "price": 19900}
        ],
        "total_quantity": 3,
        "attendees": "3명",
        "card_last4": "9716",
        "amount": 51700,
        "user_name": "",  # 영수증 실물에서 확인 불가 시 공란 처리
        "department": "",
        "notes": "정상 결제 / 수기메모: 양주시 보건소"
    }
]
