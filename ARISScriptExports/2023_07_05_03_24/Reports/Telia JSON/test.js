function getNestedChildren(arr, parent) {
  var out = []
  for (var i in arr) {
    if (arr[i].parent == parent) {
      var children = getNestedChildren(arr, arr[i].id)

      if (children.length) {
        arr[i].children = children
      }
      out.push(arr[i])
    }
  }
  return out
}


var flat = [{
    id: 1,
    title: 'hello',
    parent: 0
  },
  {
    id: 2,
    title: 'hello',
    parent: 0
  },
  {
    id: 3,
    title: 'hello',
    parent: 1
  },
  {
    id: 4,
    title: 'hello',
    parent: 3
  },
  {
    id: 5,
    title: 'hello',
    parent: 4
  },
  {
    id: 6,
    title: 'hello',
    parent: 4
  },
  {
    id: 7,
    title: 'hello',
    parent: 3
  },
  {
    id: 8,
    title: 'hello',
    parent: 2
  }
]


var nested = getNestedChildren(flat, 0)

document.write('<pre>' + JSON.stringify(nested, 0, 4) + '</pre>');