import React, { useState, useRef, useEffect } from 'react';

// 声明全局 XLSX 类型
declare global {
  interface Window {
    XLSX: any;
  }
}

interface ColoredBatch {
  row: number;
  awb: string;
  batch: string;
  boxCount: string;
}

interface BatchWithAWB {
  batch: string;
  awbs: string[];
  totalBoxCount: number;
}

const App: React.FC = () => {
  const [file, setFile] = useState<File | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string>('');
  const [batchesWithAWBs, setBatchesWithAWBs] = useState<BatchWithAWB[]>([]);
  const [details, setDetails] = useState<ColoredBatch[]>([]);
  
  const [xlsxLoaded, setXlsxLoaded] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    // 动态加载 XLSX 库
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.onload = () => {
      setXlsxLoaded(true);
      console.log('XLSX 库加载成功');
    };
    script.onerror = () => {
      setError('XLSX 库加载失败，请刷新页面重试');
    };
    document.head.appendChild(script);

    return () => {
      document.head.removeChild(script);
    };
  }, []);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      if (!selectedFile.name.endsWith('.xlsx')) {
        setError('请上传 .xlsx 格式的文件');
        return;
      }
      setFile(selectedFile);
      setError('');
      setBatchesWithAWBs([]);
      setDetails([]);
    }
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile) {
      if (!droppedFile.name.endsWith('.xlsx')) {
        setError('请上传 .xlsx 格式的文件');
        return;
      }
      setFile(droppedFile);
      setError('');
      setBatchesWithAWBs([]);
      setDetails([]);
    }
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  const processFile = async () => {
    if (!file) return;
    
    if (!xlsxLoaded || !window.XLSX) {
      setError('XLSX 库尚未加载，请稍后重试');
      return;
    }

    setIsLoading(true);
    setError('');

    try {
      const data = await file.arrayBuffer();
      const workbook = window.XLSX.read(data, {
        type: 'array',
        cellStyles: true,
        cellFormulas: true,
        cellNF: true,
        sheetStubs: true
      });

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = window.XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];

      // 查找列索引
      const headers = jsonData[0] || [];
      let awbIndex = -1;
      let batchIndex = -1;
      let boxCountIndex = -1;

      headers.forEach((header: any, index: number) => {
        const headerStr = header.toString().trim();
        if (headerStr.includes('AWB No') || headerStr.includes('提单号')) {
          awbIndex = index;
        }
        if (headerStr.includes('Collaborated Batch') || headerStr.includes('WMS协作批次')) {
          batchIndex = index;
        }
        if (headerStr.includes('大箱数') || headerStr.includes('箱数') || index === 12) { // M列是第13列，索引12
          boxCountIndex = index;
        }
      });

      if (awbIndex === -1 || batchIndex === -1) {
        setError('未找到必需的列: 提单号 AWB No 或 WMS协作批次 Collaborated Batch');
        setIsLoading(false);
        return;
      }

      // 处理合并单元格的函数
      const getMergedCellValue = (row: number, col: number): string => {
        const cellAddress = window.XLSX.utils.encode_cell({ r: row, c: col });
        const cell = sheet[cellAddress];
        
        // 如果单元格有值，直接返回
        if (cell && cell.v) {
          return cell.v.toString().trim();
        }
        
        // 检查是否在合并单元格范围内
        if (sheet['!merges']) {
          for (const merge of sheet['!merges']) {
            if (row >= merge.s.r && row <= merge.e.r &&
                col >= merge.s.c && col <= merge.e.c) {
              // 获取合并单元格的主单元格的值
              const mergedCellAddress = window.XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
              const mergedCell = sheet[mergedCellAddress];
              if (mergedCell && mergedCell.v) {
                return mergedCell.v.toString().trim();
              }
            }
          }
        }
        
        return '';
      };

      // 查找带颜色的单元格
      const coloredBatches = new Set<string>();
      const coloredDetails: ColoredBatch[] = [];
      const range = window.XLSX.utils.decode_range(sheet['!ref'] || 'A1');
      
      console.log('开始查找带颜色的AWB单元格...');
      
      // 首先统计一下样式分布，找出哪些是默认样式
      const styleCount = new Map<string, number>();
      const noStyleCount = { count: 0 };
      
      // 统计样式
      for (let row = range.s.r + 1; row <= range.e.r; row++) {
        const awbCellAddress = window.XLSX.utils.encode_cell({ r: row, c: awbIndex });
        const awbCell = sheet[awbCellAddress];
        
        if (awbCell && awbCell.v) {
          if (awbCell.s === undefined || awbCell.s === null) {
            noStyleCount.count++;
          } else {
            const styleKey = JSON.stringify(awbCell.s);
            styleCount.set(styleKey, (styleCount.get(styleKey) || 0) + 1);
          }
        }
      }
      
      // 找出最常见的样式（默认样式）
      let defaultStyle = '';
      let maxCount = 0;
      
      styleCount.forEach((count, style) => {
        if (count > maxCount) {
          maxCount = count;
          defaultStyle = style;
        }
      });
      
      console.log(`无样式单元格: ${noStyleCount.count}`);
      console.log(`最常见样式出现: ${maxCount} 次`);
      console.log(`总共 ${styleCount.size} 种样式`);
      
      // 查找带颜色的单元格
      let coloredCount = 0;
      
      for (let row = range.s.r + 1; row <= range.e.r; row++) {
        const awbCellAddress = window.XLSX.utils.encode_cell({ r: row, c: awbIndex });
        const awbCell = sheet[awbCellAddress];
        
        if (awbCell && awbCell.v) {
          const awb = awbCell.v.toString().trim();
          
          // 判断是否有颜色
          let hasColor = false;
          
          if (awbCell.s !== undefined && awbCell.s !== null) {
            const currentStyle = JSON.stringify(awbCell.s);
            
            // 如果样式不是默认样式，就认为有颜色
            if (currentStyle !== defaultStyle) {
              hasColor = true;
            }
            
            // 如果大部分单元格都没有样式，那么有样式的就是有颜色的
            if (noStyleCount.count > maxCount && awbCell.s !== undefined) {
              hasColor = true;
            }
          }
          
          if (hasColor) {
            // 获取对应的批次（处理合并单元格）
            const batch = getMergedCellValue(row, batchIndex);
            
            // 获取箱数（M列）
            let boxCount = '0';
            if (boxCountIndex !== -1) {
              const boxCountValue = getMergedCellValue(row, boxCountIndex);
              boxCount = boxCountValue || '0';
            }
            
            if (batch) {
              coloredCount++;
              coloredBatches.add(batch);
              coloredDetails.push({
                row: row + 1,
                awb: awb,
                batch: batch,
                boxCount: boxCount
              });
              
              // 输出前几个例子
              if (coloredCount <= 5) {
                console.log(`找到带颜色的单元格: 行${row + 1}, AWB=${awb}, 批次=${batch}, 箱数=${boxCount}`);
              }
            }
          }
        }
      }
      
      console.log(`总共找到 ${coloredCount} 个带颜色的AWB单元格`);
      console.log(`对应 ${coloredBatches.size} 个不同的批次`);

      // 整理数据：按批次分组AWB
      const batchAWBMap = new Map<string, { awbs: string[], totalBoxCount: number }>();
      
      coloredDetails.forEach(detail => {
        if (!batchAWBMap.has(detail.batch)) {
          batchAWBMap.set(detail.batch, { awbs: [], totalBoxCount: 0 });
        }
        const batchData = batchAWBMap.get(detail.batch)!;
        batchData.awbs.push(detail.awb);
        batchData.totalBoxCount += parseInt(detail.boxCount) || 0;
      });
      
      // 转换为数组并排序
      const batchesWithAWBsList: BatchWithAWB[] = Array.from(batchAWBMap.entries())
        .map(([batch, data]) => ({
          batch: batch,
          awbs: [...new Set(data.awbs)].sort(), // 去重并排序AWB
          totalBoxCount: data.totalBoxCount
        }))
        .sort((a, b) => a.batch.localeCompare(b.batch)); // 按批次排序

      setBatchesWithAWBs(batchesWithAWBsList);
      setDetails(coloredDetails);
      
    } catch (err) {
      setError('处理文件时出错: ' + (err as Error).message);
      console.error('错误详情:', err);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div style={{
      minHeight: '100vh',
      background: 'linear-gradient(135deg, #3498db 0%, #ffffff 50%, #ff9500 100%)',
      padding: '20px',
      margin: '0',
      position: 'fixed',
      top: '0',
      left: '0',
      right: '0',
      bottom: '0',
      overflowY: 'auto'
    }}>
      {!xlsxLoaded && (
        <div style={{
          position: 'fixed',
          top: '10px',
          right: '10px',
          background: '#ff9500',
          color: 'white',
          padding: '5px 10px',
          borderRadius: '5px',
          fontSize: '12px'
        }}>
          正在加载 XLSX 库...
        </div>
      )}
      
      <div style={{
        maxWidth: '1200px',
        margin: '0 auto',
        background: 'rgba(255, 255, 255, 0.83)',
        borderRadius: '15px',
        padding: '40px',
        boxShadow: '0 10px 30px rgba(0, 0, 0, 0.1)'
      }}>
        <h1 style={{
          textAlign: 'center',
          color: '#333',
          marginBottom: '30px',
          fontSize: '28px'
        }}>
          AWB Batch
        </h1>

        <div style={{
          display: 'flex',
          gap: '30px',
          alignItems: 'flex-start',
          flexWrap: 'wrap'
        }}>
          {/* 上传区域 */}
          <div style={{ flex: '1 1 400px', maxWidth: '700px' }}>
            <div
              onClick={() => xlsxLoaded && fileInputRef.current?.click()}
              onDrop={handleDrop}
              onDragOver={handleDragOver}
              style={{
                border: '3px dashed #3498db',
                borderRadius: '10px',
                padding: '40px 20px',
                textAlign: 'center',
                background: '#f8f9fa',
                cursor: xlsxLoaded ? 'pointer' : 'not-allowed',
                transition: 'all 0.3s ease',
                opacity: xlsxLoaded ? 1 : 0.6
              }}
            >
              <div style={{ fontSize: '48px', marginBottom: '20px' }}>📁</div>
              <div style={{ color: '#666', marginBottom: '10px' }}>
                {xlsxLoaded ? '点击或拖拽文件到此处上传' : '正在准备上传功能...'}
              </div>
              <div style={{ color: '#999', fontSize: '14px' }}>
                仅支持 .xlsx 格式
              </div>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx"
                onChange={handleFileSelect}
                style={{ display: 'none' }}
                disabled={!xlsxLoaded}
              />
            </div>

            {file && (
              <div style={{
                marginTop: '20px',
                padding: '15px',
                background: '#e8f4fd',
                borderRadius: '8px'
              }}>
                <div style={{ fontWeight: 'bold', marginBottom: '10px' }}>
                  已选择: {file.name}
                </div>
                <button
                  onClick={processFile}
                  disabled={isLoading || !xlsxLoaded}
                  style={{
                    width: '100%',
                    padding: '12px 24px',
                    background: isLoading || !xlsxLoaded ? '#95a5a6' : '#3498db',
                    color: 'white',
                    border: 'none',
                    borderRadius: '8px',
                    fontSize: '16px',
                    cursor: isLoading || !xlsxLoaded ? 'not-allowed' : 'pointer',
                    transition: 'background 0.3s ease'
                  }}
                >
                  {isLoading ? '搜索中...' : '搜索'}
                </button>
              </div>
            )}

            {error && (
              <div style={{
                marginTop: '15px',
                padding: '15px',
                background: '#ffe6e6',
                color: '#d32f2f',
                borderRadius: '8px'
              }}>
                {error}
              </div>
            )}
          </div>

          {/* 结果区域 */}
          <div style={{ flex: '2 1 400px' }}>
            <div style={{
              background: '#f8f9fa',
              borderRadius: '10px',
              padding: '20px',
              minHeight: '400px',
              border: '1px solid #e0e0e0'
            }}>
              <div style={{
                fontSize: '20px',
                fontWeight: 'bold',
                marginBottom: '20px'
              }}>
                搜索结果
              </div>

              {isLoading ? (
                <div style={{ textAlign: 'center', padding: '50px' }}>
                  <div style={{
                    display: 'inline-block',
                    width: '40px',
                    height: '40px',
                    border: '4px solid #f3f3f3',
                    borderTop: '4px solid #3498db',
                    borderRadius: '50%',
                    animation: 'spin 1s linear infinite'
                  }}/>
                </div>
              ) : batchesWithAWBs.length > 0 ? (
                <>
                  <div style={{
                    maxHeight: '300px',
                    overflowY: 'auto',
                    marginBottom: '15px'
                  }}>
                    <table style={{
                      width: '100%',
                      borderCollapse: 'collapse',
                      background: 'white',
                      borderRadius: '8px',
                      overflow: 'hidden',
                      boxShadow: '0 2px 10px rgba(0, 0, 0, 0.05)'
                    }}>
                      <thead style={{ position: 'sticky', top: 0, zIndex: 10 }}>
                        <tr>
                          <th style={{
                            background: '#3498db',
                            color: 'white',
                            padding: '15px',
                            textAlign: 'left',
                            fontWeight: 600,
                            width: '30%'
                          }}>
                            WMS协作批次
                          </th>
                          <th style={{
                            background: '#3498db',
                            color: 'white',
                            padding: '15px',
                            textAlign: 'left',
                            fontWeight: 600,
                            width: '40%'
                          }}>
                            AWB No
                          </th>
                          <th style={{
                            background: '#3498db',
                            color: 'white',
                            padding: '15px',
                            textAlign: 'center',
                            fontWeight: 600,
                            width: '25%'
                          }}>
                            大箱数
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {batchesWithAWBs.map((item, index) => (
                          <tr key={index}>
                            <td style={{
                              padding: '12px 15px',
                              borderBottom: index < batchesWithAWBs.length - 1 ? '1px solid #f0f0f0' : 'none',
                              borderRight: '1px solid #f0f0f0',
                              verticalAlign: 'top'
                            }}>
                              {item.batch}
                            </td>
                            <td style={{
                              padding: '12px 15px',
                              borderBottom: index < batchesWithAWBs.length - 1 ? '1px solid #f0f0f0' : 'none',
                              borderRight: '1px solid #f0f0f0',
                              fontSize: '13px',
                              lineHeight: '1.6'
                            }}>
                              {item.awbs.join(', ')}
                            </td>
                            <td style={{
                              padding: '12px 15px',
                              borderBottom: index < batchesWithAWBs.length - 1 ? '1px solid #f0f0f0' : 'none',
                              textAlign: 'center',
                              fontWeight: 'bold',
                              color: '#2c3e50'
                            }}>
                              {item.totalBoxCount}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div style={{
                    color: '#666',
                    fontSize: '14px',
                    marginBottom: '10px'
                  }}>
                    共找到 {batchesWithAWBs.length} 个批次，包含 {new Set(details.map(d => d.awb)).size} 个AWB
                  </div>

                  <button
                    onClick={() => {
                      // 导出为CSV
                      let csv = 'WMS协作批次,AWB No,大箱数\n';
                      batchesWithAWBs.forEach(item => {
                        csv += `"${item.batch}","${item.awbs.join(', ')}",${item.totalBoxCount}\n`;
                      });
                      
                      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                      const link = document.createElement('a');
                      link.href = URL.createObjectURL(blob);
                      link.download = `AWB批次结果_${new Date().toLocaleDateString()}.csv`;
                      link.click();
                    }}
                    style={{
                      padding: '8px 16px',
                      background: '#27ae60',
                      color: 'white',
                      border: 'none',
                      borderRadius: '5px',
                      cursor: 'pointer',
                      fontSize: '14px',
                      marginTop: '10px'
                    }}
                  >
                    导出为CSV
                  </button>

                  {/* AWB统计表格 */}
                  <div style={{
                    marginTop: '30px',
                    padding: '15px',
                    background: '#f0f8ff',
                    borderRadius: '8px'
                  }}>
                    <h3 style={{
                      fontSize: '16px',
                      marginBottom: '15px',
                      color: '#2c3e50'
                    }}>
                      AWB 统计分析
                    </h3>
                    <div style={{
                      maxHeight: '200px',
                      overflowY: 'auto'
                    }}>
                      <table style={{
                        width: '100%',
                        borderCollapse: 'collapse',
                        background: 'white',
                        borderRadius: '5px',
                        overflow: 'hidden',
                        fontSize: '13px'
                      }}>
                        <thead>
                          <tr>
                            <th style={{
                              background: '#34495e',
                              color: 'white',
                              padding: '10px',
                              textAlign: 'left',
                              position: 'sticky',
                              top: 0
                            }}>
                              AWB No
                            </th>
                            <th style={{
                              background: '#34495e',
                              color: 'white',
                              padding: '10px',
                              textAlign: 'center',
                              position: 'sticky',
                              top: 0
                            }}>
                              批次数量
                            </th>
                            <th style={{
                              background: '#34495e',
                              color: 'white',
                              padding: '10px',
                              textAlign: 'center',
                              position: 'sticky',
                              top: 0
                            }}>
                              总箱数
                            </th>
                          </tr>
                        </thead>
                        <tbody>
                          {(() => {
                            // 统计每个AWB对应的批次数量和箱数
                            const awbStats = new Map<string, { count: number, totalBoxes: number }>();
                            details.forEach(detail => {
                              if (!awbStats.has(detail.awb)) {
                                awbStats.set(detail.awb, { count: 0, totalBoxes: 0 });
                              }
                              const stats = awbStats.get(detail.awb)!;
                              stats.count += 1;
                              stats.totalBoxes += parseInt(detail.boxCount) || 0;
                            });
                            
                            // 转换为数组并排序
                            const sortedStats = Array.from(awbStats.entries())
                              .sort((a, b) => b[1].count - a[1].count); // 按批次数量降序排序
                            
                            return sortedStats.map(([awb, stats]) => (
                              <tr key={awb}>
                                <td style={{
                                  padding: '8px 10px',
                                  borderBottom: '1px solid #ecf0f1'
                                }}>
                                  {awb}
                                </td>
                                <td style={{
                                  padding: '8px 10px',
                                  borderBottom: '1px solid #ecf0f1',
                                  textAlign: 'center',
                                  fontWeight: stats.count > 1 ? 'bold' : 'normal',
                                  color: stats.count > 1 ? '#e74c3c' : '#2c3e50'
                                }}>
                                  {stats.count}
                                </td>
                                <td style={{
                                  padding: '8px 10px',
                                  borderBottom: '1px solid #ecf0f1',
                                  textAlign: 'center',
                                  fontWeight: 'bold',
                                  color: '#27ae60'
                                }}>
                                  {stats.totalBoxes}
                                </td>
                              </tr>
                            ));
                          })()}
                        </tbody>
                      </table>
                    </div>
                    <div style={{
                      marginTop: '10px',
                      fontSize: '12px',
                      color: '#7f8c8d'
                    }}>
                      共 {new Set(details.map(d => d.awb)).size} 个不同的AWB
                    </div>
                  </div>
                </>
              ) : (
                <div style={{
                  textAlign: 'center',
                  color: '#999',
                  padding: '50px',
                  fontSize: '16px'
                }}>
                  请上传文件并点击搜索
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      <style>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
    </div>
  );
};

export default App;