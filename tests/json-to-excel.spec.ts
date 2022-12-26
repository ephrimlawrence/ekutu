import { test } from '@japa/runner';
import { JsonToExcel } from "../src/index";

const json = [
  {
    id: 1,
    name: 'John',
    age: 25,
    address: {
      street: 'Main Street',
      city: 'London',
      country: 'UK',
      workExperience: {
        startDate: '2016-01-01',
        endDate: '2016-12-31',
      }
    }
  },
  {
    id: 2,
    name: 'Jane',
    age: 30,
    address: {
      street: 'Main Street',
      city: 'London',
      country: 'UK',
    }
  },
  {
    id: 3,
    name: 'Jack',
  }
];

test.group('JsonToExcel', () => {
  test('basic export', ({ assert }) => {
    const toExcel = new JsonToExcel();
    const workbook = toExcel.export({ data: json });

    workbook.xlsx.writeFile("test.xlsx");
    // Test logic goes here
    assert.equal(2 + 2, 4)
  })
})
