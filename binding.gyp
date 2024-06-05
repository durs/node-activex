{
  'targets': [
    {
      'target_name': 'node_activex',
      'conditions': [
        ['OS=="win"', {
          'sources': [
            'src/main.cpp',
            'src/utils.cpp',
            'src/disp.cpp'
          ],
          'libraries': [],
          'dependencies': []
        }]
      ]
    }
  ]
}