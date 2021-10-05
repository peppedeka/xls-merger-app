import 'jspdf-autotable';

import {HttpClient} from '@angular/common/http';
import {Component, EventEmitter} from '@angular/core';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import {BehaviorSubject, forkJoin, from, Observable} from 'rxjs';
import {debounceTime, map, switchMap} from 'rxjs/operators';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  myFiles: any[] = [];
  private _workBook: any = null;
  private _readerEvt: EventEmitter<Observable<any>> =
      new EventEmitter<Observable<any>>();
  rows$: BehaviorSubject<any[]> = new BehaviorSubject<any[]>([]);
  constructor(private http: HttpClient) {
    this._readerEvt
        .pipe(
            debounceTime(300), switchMap((d) => d), map((dataList: any[]) => {
              let jsonDataList: any[] = [];
              dataList.forEach(data => {
                this._workBook =
                    XLSX.read(data, {type: 'binary', cellStyles: true});
                jsonDataList.push(
                    XLSX.utils.sheet_to_json(this._workBook.Sheets['Table 1']));
              });
              return jsonDataList;
            }),
            map((list: {[key: string]: string}[][]) => {
              let res: {
                [iva: number]: {
                  [articolo: number]: {
                    articolo: number,
                    descrizione: string,
                    quantita: number,
                    valoreUnitario: number,
                    valoreTotale: number,
                  }
                }
              } = {4: {}, 10: {}};
              for (let i = 0; i < list.length; i++) {
                const xlsxData = list[i];
                let iva = 4;

                for (let j = 5; j < xlsxData.length; j++) {
                  const xlsRow = Object.values(xlsxData[j]);
                  if (xlsRow.length === 5) {
                    const articolo: number = +xlsRow[0];
                    let descrizione: string = xlsRow[1];
                    const quantita: number = +xlsRow[2];
                    const valoreUnitario: number = +xlsRow[3];

                    if (res[iva][articolo] != null) {
                      res[iva][articolo].quantita += quantita;
                      res[iva][articolo].valoreTotale =
                          res[iva][articolo].quantita *
                          res[iva][articolo].valoreUnitario;
                    } else {
                      const valoreTotale: number = quantita * valoreUnitario;
                      if (descrizione.includes('Totale al       4,00 %')) {
                        iva = 10;
                        descrizione =
                            descrizione.replace('Totale al       4,00 %', '');
                      }
                      res[iva][articolo] = {
                        articolo,
                        descrizione,
                        quantita,
                        valoreUnitario,
                        valoreTotale,
                      }
                    }
                  }
                }
              }

              return res;
            }),
            map((objectData) => {
              const emptyRow = {
                articolo: '',
                descrizione: '',
                quantita: '',
                valoreUnitario: '',
                valoreTotale: ''
              };
              const quantitaTotale4 =
                  Object.values(objectData[4])
                      .map((v) => v.quantita)
                      .reduce((prev, current) => prev + current);
              const quantitaTotale10 =
                  Object.values(objectData[10])
                      .map((v) => v.quantita)
                      .reduce((prev, current) => prev + current);
              const valoreTotale4 =
                  Math.round(Object.values(objectData[4])
                                 .map((v) => v.valoreTotale)
                                 .reduce((prev, current) => prev + current));
              ;
              const valoreTotale10 =
                  Math.round(Object.values(objectData[10])
                                 .map((v) => v.valoreTotale)
                                 .reduce((prev, current) => prev + current));
              const jsonData: any[] =
                  [
                    ...Object.values(objectData[4]), {
                      articolo: '',
                      descrizione: 'Totale al       4,00 %',
                      quantita: quantitaTotale4,
                      valoreUnitario: '',
                      valoreTotale: valoreTotale4
                    },
                    emptyRow, ...Object.values(objectData[10]), {
                      articolo: '',
                      descrizione: 'Totale al       10,00 %',
                      quantita: quantitaTotale10,
                      valoreUnitario: '',
                      valoreTotale: valoreTotale10
                    },
                    emptyRow, {
                      articolo: 'TOTALE',
                      descrizione: '',
                      quantita: Math.round(quantitaTotale4 + quantitaTotale10),
                      valoreUnitario: '',
                      valoreTotale: Math.round(valoreTotale4 + valoreTotale10)
                    }
                  ]

                  return jsonData;
            }))
        .subscribe((rows) => {
          this.rows$.next(rows);
        })
  }
  readFileAsText(file: File): Observable<any> {
    const promi = new Promise(function(resolve, reject) {
      let fr = new FileReader();

      fr.onload = function() {
        resolve(fr.result);
      };

      fr.onerror = function() {
        reject(fr);
      };

      fr.readAsBinaryString(file);
    });
    return from(promi);
  }
  onFileChange(ev: any) {
    const obsList: Observable<any>[] = [];
    const fileNum = (ev.target.files as FileList).length;
    for (let i = 0; i < fileNum; i++) {
      const file = ev.target.files[i];
      obsList.push(this.readFileAsText(file));
    }
    this._readerEvt.emit(forkJoin(obsList))
  }
  download(): void {
    let doc = new jsPDF();
    autoTable(doc, {html: '#table', theme: 'striped'});
    (doc as any)
        .save('table.pdf')

            (doc as any)
        .autoTable({html: '#table'});
    let datauri = (doc as any).output('datauri', 'test.pdf');
    var iframe = '<iframe width=\'100%\' height=\'100%\' src=\'' + datauri +
        '\'></iframe>'
    var x: any = window.open();
    x.document.open();
    x.document.write(iframe);
    x.document.close();
  }
}
