let data = [
  {
    age: 24,
    sex: "male",
    name: "john",
    signed_up: false,
    tables: [
      { name: "haha", age: 35 },
      { name: "wowo", age: 40 }
    ]
  },
  {
    age: 30,
    sex: "male",
    name: "mike",
    signed_up: true,
    tables: [
      { name: "haha", age: 35 },
      { name: "wowo", age: 40 }
    ]
  }
]

$data | to xlsx test.xlsx
