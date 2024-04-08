import '@univerjs/design/lib/index.css';
import '@univerjs/ui/lib/index.css';
import '@univerjs/sheets-ui/lib/index.css';
import '@univerjs/sheets-formula/lib/index.css';
import './index.css';
import {
  Univer,
} from "@univerjs/core";
import { defaultTheme } from '@univerjs/design';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
import { UniverUIPlugin } from '@univerjs/ui';
import { forwardRef, useEffect, useRef, useState } from 'react';
import ImportExcelButtonPlugin from '../../plugins/ImportExcelButton';
import ExportExcelButtonPlugin from '../../plugins/ExportExcelButton';
import { MY_DATA } from '../../assets/my-data'


// eslint-disable-next-line react/display-name
const UniverSheet = forwardRef(() => {
  const univerRef = useRef(null);
  const workbookRef = useRef(null);
  const containerRef = useRef(null);
  const [univeData, setUniveData] = useState(MY_DATA);

  const handleImportExcel = (data) => {
    // 更新组件状态以触发刷新
    setUniveData(data);
  }

  useEffect(() => {
    // 在组件加载后设置回调函数
    ImportExcelButtonPlugin.setOnImportExcelCallback(handleImportExcel);
  }, []);

  /**
   * Initialize univer instance and workbook instance
   * @param data {IWorkbookData} document see https://univer.work/api/core/interfaces/IWorkbookData.html
   */
const data = {
  "id": "workbook-01",
  "sheetOrder": [
      "s1",
      "ss2",
      "Sheet3"
  ],
  "name": "universheet",
  "appVersion": "3.0.0-alpha",
  "locale": "zhCN",
  "styles": {},
  "sheets": {
      "s1": {
          "id": "s1",
          "name": "s1",
          "tabColor": "white",
          "hidden": 0,
          "freeze": {
              "xSplit": 0,
              "ySplit": 0,
              "startRow": 0,
              "startColumn": 0
          },
          "rowCount": 1000,
          "columnCount": 100,
          "zoomRatio": 1,
          "scrollTop": 200,
          "scrollLeft": 100,
          "defaultColumnWidth": 93,
          "defaultRowHeight": 27,
          "mergeData": [
              {
                  "startRow": 3,
                  "startColumn": 2,
                  "endRow": 3,
                  "endColumn": 3
              },
              {
                  "startRow": 4,
                  "startColumn": 4,
                  "endRow": 4,
                  "endColumn": 9
              },
              {
                  "startRow": 6,
                  "startColumn": 6,
                  "endRow": 13,
                  "endColumn": 6
              },
              {
                  "startRow": 10,
                  "startColumn": 3,
                  "endRow": 11,
                  "endColumn": 3
              }
          ],
          "cellData": {
              "0": {
                  "0": {
                      "v": "序号"
                  },
                  "1": {
                      "v": "销售数量"
                  },
                  "2": {},
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "1": {
                  "0": {
                      "v": 1
                  },
                  "1": {
                      "v": 1000
                  },
                  "2": {},
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "2": {
                  "0": {
                      "v": 2
                  },
                  "1": {
                      "v": 100
                  },
                  "2": {},
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "3": {
                  "0": {
                      "v": 3
                  },
                  "1": {
                      "v": 520
                  },
                  "2": {
                      "v": 12
                  },
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "4": {
                  "0": {
                      "v": 4
                  },
                  "1": {
                      "v": 512
                  },
                  "2": {},
                  "3": {},
                  "4": {
                      "v": "行单元格"
                  },
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "5": {
                  "0": {
                      "v": 5
                  },
                  "1": {
                      "v": 1111
                  },
                  "2": {},
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "6": {
                  "0": {},
                  "1": {},
                  "2": {},
                  "3": {
                      "v": "批注"
                  },
                  "4": {},
                  "5": {},
                  "6": {
                      "v": "列单元格"
                  },
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "7": {
                  "0": {
                      "v": 7
                  },
                  "1": {},
                  "2": {},
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "8": {
                  "0": {},
                  "1": {
                      "v": 200
                  },
                  "2": {},
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "9": {
                  "0": {
                      "v": 9
                  },
                  "1": {
                      "v": 300
                  },
                  "2": {},
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "10": {
                  "0": {},
                  "1": {},
                  "2": {},
                  "3": {
                      "v": 24
                  },
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "11": {
                  "0": {},
                  "1": {},
                  "2": {},
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "12": {
                  "0": {},
                  "1": {},
                  "2": {},
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              },
              "13": {
                  "0": {},
                  "1": {},
                  "2": {},
                  "3": {},
                  "4": {},
                  "5": {},
                  "6": {},
                  "7": {},
                  "8": {},
                  "9": {}
              }
          },
          "rowData": {
              "0": {
                  "ah": 27
              },
              "1": {},
              "2": {},
              "3": {},
              "4": {},
              "5": {},
              "6": {},
              "7": {},
              "8": {},
              "9": {},
              "10": {},
              "11": {},
              "12": {},
              "13": {}
          },
          "columnData": {
              "0": {},
              "1": {},
              "2": {},
              "3": {},
              "4": {},
              "5": {},
              "6": {},
              "7": {},
              "8": {},
              "9": {}
          },
          "rowHeader": {
              "width": 46,
              "hidden": 0
          },
          "columnHeader": {
              "height": 20,
              "hidden": 0
          },
          "showGridlines": 1,
          "selections": [
              "A2"
          ],
          "rightToLeft": 0
      },
      "ss2": {
          "id": "ss2",
          "name": "ss2",
          "tabColor": "white",
          "hidden": 0,
          "freeze": {
              "xSplit": 0,
              "ySplit": 0,
              "startRow": 0,
              "startColumn": 0
          },
          "rowCount": 1000,
          "columnCount": 100,
          "zoomRatio": 1,
          "scrollTop": 200,
          "scrollLeft": 100,
          "defaultColumnWidth": 93,
          "defaultRowHeight": 27,
          "mergeData": [],
          "cellData": {
              "0": {
                  "0": {
                      "v": "hello"
                  }
              }
          },
          "rowData": {
              "0": {}
          },
          "columnData": {
              "0": {}
          },
          "rowHeader": {
              "width": 46,
              "hidden": 0
          },
          "columnHeader": {
              "height": 20,
              "hidden": 0
          },
          "showGridlines": 1,
          "selections": [
              "A2"
          ],
          "rightToLeft": 0
      },
      "Sheet3": {
          "id": "Sheet3",
          "name": "Sheet3",
          "tabColor": "white",
          "hidden": 0,
          "freeze": {
              "xSplit": 0,
              "ySplit": 0,
              "startRow": 0,
              "startColumn": 0
          },
          "rowCount": 1000,
          "columnCount": 100,
          "zoomRatio": 1,
          "scrollTop": 200,
          "scrollLeft": 100,
          "defaultColumnWidth": 93,
          "defaultRowHeight": 27,
          "mergeData": [],
          "cellData": {
              "0": {
                  "0": {
                      "v": "dada"
                  }
              }
          },
          "rowData": {
              "0": {}
          },
          "columnData": {
              "0": {}
          },
          "rowHeader": {
              "width": 46,
              "hidden": 0
          },
          "columnHeader": {
              "height": 20,
              "hidden": 0
          },
          "showGridlines": 1,
          "selections": [
              "A2"
          ],
          "rightToLeft": 0
      }
  },
  "resources": []
}
const workdata = 
{
    "id": "workbook-01",
    "sheetOrder": [
        "s1",
        "ss2",
        "Sheet3"
    ],
    "name": "universheet",
    "appVersion": "3.0.0-alpha",
    "locale": "zhCN",
    "styles": {},
    "sheets": {
        "s1": {
            "id": "1",
            "name": "s1",
            "tabColor": "white",
            "hidden": 0,
            "freeze": {
                "xSplit": 0,
                "ySplit": 0,
                "startRow": 0,
                "startColumn": 0
            },
            "rowCount": 1000,
            "columnCount": 100,
            "zoomRatio": 1,
            "scrollTop": 200,
            "scrollLeft": 100,
            "defaultColumnWidth": 93,
            "defaultRowHeight": 27,
            "mergeData": [
                {
                    "startRow": 3,
                    "startColumn": 2,
                    "endRow": 3,
                    "endColumn": 3
                },
                {
                    "startRow": 4,
                    "startColumn": 4,
                    "endRow": 4,
                    "endColumn": 9
                },
                {
                    "startRow": 6,
                    "startColumn": 6,
                    "endRow": 13,
                    "endColumn": 6
                },
                {
                    "startRow": 10,
                    "startColumn": 3,
                    "endRow": 11,
                    "endColumn": 3
                }
            ],
            "cellData": {
                "0": {
                    "0": {
                        "v": "序号"
                    },
                    "1": {
                        "v": "销售数量"
                    },
                    "2": {},
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "1": {
                    "0": {
                        "v": 1
                    },
                    "1": {
                        "v": 1000
                    },
                    "2": {},
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "2": {
                    "0": {
                        "v": 2
                    },
                    "1": {
                        "v": 100
                    },
                    "2": {},
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "3": {
                    "0": {
                        "v": 3
                    },
                    "1": {
                        "v": 520
                    },
                    "2": {
                        "v": 12
                    },
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "4": {
                    "0": {
                        "v": 4
                    },
                    "1": {
                        "v": 512
                    },
                    "2": {},
                    "3": {},
                    "4": {
                        "v": "行单元格"
                    },
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "5": {
                    "0": {
                        "v": 5
                    },
                    "1": {
                        "v": 1111
                    },
                    "2": {},
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "6": {
                    "0": {},
                    "1": {},
                    "2": {},
                    "3": {
                        "v": "批注"
                    },
                    "4": {},
                    "5": {},
                    "6": {
                        "v": "列单元格"
                    },
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "7": {
                    "0": {
                        "v": 7
                    },
                    "1": {},
                    "2": {},
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "8": {
                    "0": {},
                    "1": {
                        "v": 200
                    },
                    "2": {},
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "9": {
                    "0": {
                        "v": 9
                    },
                    "1": {
                        "v": 300
                    },
                    "2": {},
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "10": {
                    "0": {},
                    "1": {},
                    "2": {},
                    "3": {
                        "v": 24
                    },
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "11": {
                    "0": {},
                    "1": {},
                    "2": {},
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "12": {
                    "0": {},
                    "1": {},
                    "2": {},
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                },
                "13": {
                    "0": {},
                    "1": {},
                    "2": {},
                    "3": {},
                    "4": {},
                    "5": {},
                    "6": {},
                    "7": {},
                    "8": {},
                    "9": {}
                }
            },
            "rowData": {
                "0": {},
                "1": {},
                "2": {},
                "3": {},
                "4": {},
                "5": {},
                "6": {},
                "7": {},
                "8": {},
                "9": {},
                "10": {},
                "11": {},
                "12": {},
                "13": {}
            },
            "columnData": {
                "0": {},
                "1": {},
                "2": {},
                "3": {},
                "4": {},
                "5": {},
                "6": {},
                "7": {},
                "8": {},
                "9": {}
            },
            "rowHeader": {
                "width": 46,
                "hidden": 0
            },
            "columnHeader": {
                "height": 20,
                "hidden": 0
            },
            "showGridlines": 1,
            "selections": [
                "A2"
            ],
            "rightToLeft": 0
        },
        "ss2": {
            "id": "2",
            "name": "ss2",
            "tabColor": "white",
            "hidden": 0,
            "freeze": {
                "xSplit": 0,
                "ySplit": 0,
                "startRow": 0,
                "startColumn": 0
            },
            "rowCount": 1000,
            "columnCount": 100,
            "zoomRatio": 1,
            "scrollTop": 200,
            "scrollLeft": 100,
            "defaultColumnWidth": 93,
            "defaultRowHeight": 27,
            "mergeData": [],
            "cellData": {
                "0": {
                    "0": {
                        "v": "hello"
                    }
                }
            },
            "rowData": {
                "0": {}
            },
            "columnData": {
                "0": {}
            },
            "rowHeader": {
                "width": 46,
                "hidden": 0
            },
            "columnHeader": {
                "height": 20,
                "hidden": 0
            },
            "showGridlines": 1,
            "selections": [
                "A2"
            ],
            "rightToLeft": 0
        },
        "Sheet3": {
            "id": "3",
            "name": "Sheet3",
            "tabColor": "white",
            "hidden": 0,
            "freeze": {
                "xSplit": 0,
                "ySplit": 0,
                "startRow": 0,
                "startColumn": 0
            },
            "rowCount": 1000,
            "columnCount": 100,
            "zoomRatio": 1,
            "scrollTop": 200,
            "scrollLeft": 100,
            "defaultColumnWidth": 93,
            "defaultRowHeight": 27,
            "mergeData": [],
            "cellData": {
                "0": {
                    "0": {
                        "v": "dada"
                    }
                }
            },
            "rowData": {
                "0": {}
            },
            "columnData": {
                "0": {}
            },
            "rowHeader": {
                "width": 46,
                "hidden": 0
            },
            "columnHeader": {
                "height": 20,
                "hidden": 0
            },
            "showGridlines": 1,
            "selections": [
                "A2"
            ],
            "rightToLeft": 0
        }
    },
    "resources": []
}

// You can now use the `workbookData` variable to access the workbook information as needed.

  const init = () => {
    if (!containerRef.current) {
      throw Error('container not initialized');
    }
    const univer = new Univer({
      theme: defaultTheme,
    });
    univerRef.current = univer;

    // core plugins
    univer.registerPlugin(UniverRenderEnginePlugin);
    univer.registerPlugin(UniverFormulaEnginePlugin);
    univer.registerPlugin(UniverUIPlugin, {
      container: containerRef.current,
      header: true,
      toolbar: true,
      footer: true,
    });

    // doc plugins
    univer.registerPlugin(UniverDocsPlugin, {
      hasScroll: false,
    });
    univer.registerPlugin(UniverDocsUIPlugin);

    // sheet plugins
    univer.registerPlugin(UniverSheetsPlugin);
    univer.registerPlugin(UniverSheetsUIPlugin);
    univer.registerPlugin(UniverSheetsFormulaPlugin);
    univer.registerPlugin(ImportExcelButtonPlugin);
    univer.registerPlugin(ExportExcelButtonPlugin);
    // create workbook instance
    workbookRef.current = univer.createUniverSheet(workdata);
  };

  /**
   * Destroy univer instance and workbook instance
   */
  const destroyUniver = () => {
    univerRef.current?.dispose();
    univerRef.current = null;
    workbookRef.current = null;
  };

  useEffect(() => {
    console.log('init 开始')
    console.log(univeData)
    init();
    return () => {
      destroyUniver();
    };
  }, [univeData]);

  return <div ref={containerRef} className="univer-container" />;
});

export default UniverSheet;
