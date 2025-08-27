let data = [
  {
    age: 24,
    sex: "male",
    name: "john",
    signed_up: false,
    payload: [
      { nick_name: "haha", weight: 100 },
      { nick_name: "wowo", weight: 40 }
    ]
  },
  {
    age: 30,
    sex: "male",
    name: "mike",
    signed_up: true,
    payload: [
      { name: "haha", age: 35 },
      { name: "wowo", age: 40 }
    ]
  }
]

$data | to xlsx test.xlsx
