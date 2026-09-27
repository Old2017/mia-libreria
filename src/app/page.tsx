'use client';

import React, { useEffect, useRef, useState } from 'react';
import {
  Home,
  BookOpen,
  Users,
  Settings,
  Plus,
  Search,
  Scan,
  Trash2,
  Edit,
  Download,
  Upload,
  CheckCircle2,
  Book,
  ChevronRight,
  Star,
  X,
  FileSpreadsheet,
  FileText,
  Calendar,
  GripVertical,
  Layers,
  Sparkles,
} from 'lucide-react';

import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

import {
  BrowserMultiFormatReader,
  type IScannerControls,
} from '@zxing/browser';

import {
  BarcodeFormat,
  DecodeHintType,
} from '@zxing/library';

/* =========================================================
   TIPI
========================================================= */

export interface BookItem {
  id: string;
  title: string;
  author: string;
  publishCountry?: string;
  coverUrl?: string;
  publisher?: string;
  publishYear?: string;
  pages?: number;
  genre?: string;
  seriesTag?: string;
  volume?: string;
  isClassic?: boolean;
  format: 'cartaceo' | 'ebook' | 'ebook_and_paper';
  isRead: boolean;
  readMonth?: string;
  readYear?: number;
  readMonthYear?: string;
  rating?: number;
  notes?: string;
  isbn?: string;
  createdAt: number;
}

const MONTHS = [
  'Gennaio',
  'Febbraio',
  'Marzo',
  'Aprile',
  'Maggio',
  'Giugno',
  'Luglio',
  'Agosto',
  'Settembre',
  'Ottobre',
  'Novembre',
  'Dicembre',
];

const STORAGE_KEY = 'ios_library_books_v5';

/* =========================================================
   COMPONENTE
========================================================= */

export default function LibraryApp() {
  /* =======================================================
     NAVIGAZIONE
  ======================================================= */

  const [activeTab, setActiveTab] = useState<
    'home' | 'read' | 'authors' | 'settings'
  >('home');

  const [books, setBooks] = useState<BookItem[]>([]);
  const [isLoaded, setIsLoaded] = useState(false);

  /* =======================================================
     MODALI
  ======================================================= */

  const [isAddModalOpen, setIsAddModalOpen] = useState(false);

  const [selectedBookDetail, setSelectedBookDetail] =
    useState<BookItem | null>(null);

  const [selectedAuthor, setSelectedAuthor] =
    useState<string | null>(null);

  const [homeSubView, setHomeSubView] = useState<
    'none' | 'classics' | 'genres'
  >('none');

  const [selectedGenreHome, setSelectedGenreHome] =
    useState<string | null>(null);

  /* =======================================================
     DRAG & DROP
  ======================================================= */

  const [draggedIndex, setDraggedIndex] =
    useState<number | null>(null);

  /* =======================================================
     ISBN
  ======================================================= */

  const [isbnInput, setIsbnInput] = useState('');
  const [isSearchingIsbn, setIsSearchingIsbn] =
    useState(false);

  /* =======================================================
     SCANNER
  ======================================================= */

  const [isScannerOpen, setIsScannerOpen] =
    useState(false);

  const [isScanning, setIsScanning] =
    useState(false);

  const [scannerError, setScannerError] =
    useState('');

  const [scannerStatus, setScannerStatus] =
    useState(
      'Posiziona il codice a barre del libro nel riquadro'
    );

  const scannerVideoRef =
    useRef<HTMLVideoElement | null>(null);

  const scannerControlsRef =
    useRef<IScannerControls | null>(null);

  const scannerReaderRef =
    useRef<BrowserMultiFormatReader | null>(null);

  const scanLockedRef =
    useRef(false);

  /* =======================================================
     FILTRI
  ======================================================= */

  const [filterGenre, setFilterGenre] =
    useState('all');

  const [filterFormat, setFilterFormat] =
    useState('all');

  const [filterYear, setFilterYear] =
    useState('all');

  const [searchQuery, setSearchQuery] =
    useState('');

  /* =======================================================
     DATE
  ======================================================= */

  const currentYearNum =
    new Date().getFullYear();

  const startYear = 2023;

  const yearsList = Array.from(
    {
      length: Math.max(
        1,
        currentYearNum - startYear + 1
      ),
    },
    (_, i) => currentYearNum - i
  );

  /* =======================================================
     FORM
  ======================================================= */

  const [formData, setFormData] =
    useState<Partial<BookItem>>({
      title: '',
      author: '',
      publishCountry: '',
      coverUrl: '',
      publisher: '',
      publishYear: '',
      pages: undefined,
      genre: '',
      seriesTag: '',
      volume: '',
      isClassic: false,
      format: 'cartaceo',
      isRead: false,
      readMonth: 'Gennaio',
      readYear: currentYearNum,
      rating: 5,
      notes: '',
      isbn: '',
    });

  /* =======================================================
     CARICAMENTO
  ======================================================= */

  useEffect(() => {
    try {
      const saved =
        localStorage.getItem(STORAGE_KEY);

      if (saved) {
        const parsed = JSON.parse(saved);

        if (Array.isArray(parsed)) {
          setBooks(parsed);
        }
      }
    } catch (error) {
      console.error(
        'Errore caricamento storage:',
        error
      );
    }

    setIsLoaded(true);
  }, []);

  /* =======================================================
     SALVATAGGIO
  ======================================================= */

  useEffect(() => {
    if (!isLoaded) return;

    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify(books)
      );
    } catch (error) {
      console.error(
        'Errore salvataggio storage:',
        error
      );
    }
  }, [books, isLoaded]);

  /* =======================================================
     NORMALIZZAZIONE ISBN
  ======================================================= */

  const normalizeISBN = (
    value: string
  ): string => {
    return value
      .replace(/[^0-9Xx]/g, '')
      .toUpperCase();
  };

  const isbn10To13 = (
    isbn10: string
  ): string | null => {
    const clean =
      normalizeISBN(isbn10);

    if (
      clean.length !== 10 ||
      !/^\d{9}[\dX]$/.test(clean)
    ) {
      return null;
    }

    const body =
      '978' + clean.substring(0, 9);

    let sum = 0;

    for (let i = 0; i < body.length; i++) {
      const digit =
        Number(body[i]);

      sum +=
        i % 2 === 0
          ? digit
          : digit * 3;
    }

    const check =
      (10 - (sum % 10)) % 10;

    return body + check;
  };

  const getISBNFromBarcode = (
    rawValue: string
  ): string | null => {
    const value =
      normalizeISBN(rawValue);

    if (value.length === 13) {
      if (
        value.startsWith('978') ||
        value.startsWith('979')
      ) {
        return value;
      }
    }

    if (value.length === 10) {
      return isbn10To13(value);
    }

    /*
     * Alcuni scanner possono restituire
     * dati aggiuntivi.
     */
    if (value.length > 13) {
      const first13 =
        value.substring(0, 13);

      if (
        first13.startsWith('978') ||
        first13.startsWith('979')
      ) {
        return first13;
      }
    }

    return null;
  };

  /* =======================================================
     STOP SCANNER
  ======================================================= */

  const stopScanner = () => {
    try {
      scannerControlsRef.current?.stop();
    } catch (error) {
      console.error(
        'Errore chiusura scanner:',
        error
      );
    }

    scannerControlsRef.current = null;

    const video =
      scannerVideoRef.current;

    if (video?.srcObject) {
      const stream =
        video.srcObject as MediaStream;

      stream
        .getTracks()
        .forEach((track) => {
          track.stop();
        });

      video.srcObject = null;
    }

    try {
      scannerReaderRef.current?.reset();
    } catch {
      // niente
    }

    scannerReaderRef.current = null;

    scanLockedRef.current = false;

    setIsScanning(false);
  };

  /* =======================================================
     APERTURA SCANNER
  ======================================================= */

  const openScanner = () => {
    setScannerError('');

    setScannerStatus(
      'Posiziona il codice a barre del libro nel riquadro'
    );

    scanLockedRef.current = false;

    setIsScannerOpen(true);
  };

  /* =======================================================
     CHIUSURA SCANNER
  ======================================================= */

  const closeScanner = () => {
    stopScanner();

    setIsScannerOpen(false);

    setScannerError('');
  };

  /* =======================================================
     AVVIO SCANNER
  ======================================================= */

  useEffect(() => {
    if (!isScannerOpen) return;

    let cancelled = false;

    const startScanner =
      async () => {
        try {
          setScannerError('');
          setIsScanning(false);

          if (
            !navigator.mediaDevices ||
            !navigator.mediaDevices.getUserMedia
          ) {
            throw new Error(
              'La fotocamera non è disponibile. Apri il sito tramite HTTPS.'
            );
          }

          /*
           * Limitiamo la ricerca ai codici più comuni
           * sui libri.
           *
           * ISBN-13 -> EAN-13
           */
          const hints = new Map();

          hints.set(
            DecodeHintType.POSSIBLE_FORMATS,
            [
              BarcodeFormat.EAN_13,
              BarcodeFormat.EAN_8,
              BarcodeFormat.UPC_A,
              BarcodeFormat.UPC_E,
              BarcodeFormat.CODE_128,
            ]
          );

          hints.set(
            DecodeHintType.TRY_HARDER,
            true
          );

          const reader =
            new BrowserMultiFormatReader(
              hints
            );

          scannerReaderRef.current =
            reader;

          /*
           * Richiediamo esplicitamente la
           * fotocamera posteriore.
           */
          const stream =
            await navigator.mediaDevices.getUserMedia(
              {
                video: {
                  facingMode: {
                    ideal: 'environment',
                  },
                  width: {
                    ideal: 1920,
                  },
                  height: {
                    ideal: 1080,
                  },
                  frameRate: {
                    ideal: 30,
                  },
                },
                audio: false,
              }
            );

          if (cancelled) {
            stream
              .getTracks()
              .forEach((track) =>
                track.stop()
              );

            return;
          }

          const video =
            scannerVideoRef.current;

          if (!video) {
            stream
              .getTracks()
              .forEach((track) =>
                track.stop()
              );

            throw new Error(
              'Video della fotocamera non disponibile.'
            );
          }

          video.srcObject = stream;

          video.muted = true;

          video.playsInline = true;

          video.setAttribute(
            'playsinline',
            'true'
          );

          video.setAttribute(
            'webkit-playsinline',
            'true'
          );

          await video.play();

          if (cancelled) {
            stopScanner();
            return;
          }

          setIsScanning(true);

          /*
           * ZXing continua a leggere il video.
           */
          const controls =
            await reader.decodeFromVideoDevice(
              undefined,
              video,
              (result, error) => {
                if (cancelled) return;

                if (!result) return;

                if (
                  scanLockedRef.current
                ) {
                  return;
                }

                const raw =
                  result.getText();

                console.log(
                  'Barcode rilevato:',
                  raw
                );

                const isbn =
                  getISBNFromBarcode(
                    raw
                  );

                /*
                 * Potrebbe essere un EAN che
                 * non rappresenta un ISBN.
                 */
                if (!isbn) {
                  setScannerStatus(
                    'Codice rilevato. Cerca il codice ISBN-13 sul retro del libro.'
                  );

                  return;
                }

                /*
                 * Blocchiamo immediatamente
                 * ulteriori letture.
                 */
                scanLockedRef.current =
                  true;

                setScannerStatus(
                  `ISBN rilevato: ${isbn}`
                );

                setIsbnInput(isbn);

                stopScanner();

                setIsScannerOpen(false);

                /*
                 * Facciamo partire Google Books
                 * dopo aver chiuso la camera.
                 */
                window.setTimeout(() => {
                  handleSearchBookByISBN(
                    isbn
                  );
                }, 150);
              }
            );

          scannerControlsRef.current =
            controls;
        } catch (error: any) {
          console.error(
            'Errore avvio scanner:',
            error
          );

          if (cancelled) return;

          setIsScanning(false);

          let message =
            'Impossibile avviare la fotocamera.';

          if (
            error?.name ===
            'NotAllowedError'
          ) {
            message =
              'Accesso alla fotocamera negato. Controlla Impostazioni > Safari > Fotocamera.';
          } else if (
            error?.name ===
            'NotFoundError'
          ) {
            message =
              'Nessuna fotocamera disponibile.';
          } else if (
            error?.name ===
            'NotReadableError'
          ) {
            message =
              'La fotocamera è già utilizzata da un’altra applicazione.';
          } else if (
            error?.name ===
            'SecurityError'
          ) {
            message =
              'Safari ha bloccato la fotocamera. Usa HTTPS.';
          } else if (
            error?.message
          ) {
            message =
              error.message;
          }

          setScannerError(message);
        }
      };

    startScanner();

    return () => {
      cancelled = true;
      stopScanner();
    };
  }, [isScannerOpen]);

  /* =======================================================
     RICERCA GOOGLE BOOKS
  ======================================================= */

  const handleSearchBookByISBN =
    async (
      codeToSearch?: string
    ) => {
      const query =
        normalizeISBN(
          codeToSearch ||
            isbnInput
        );

      if (!query) {
        alert(
          'Inserisci un ISBN.'
        );
        return;
      }

      setIsSearchingIsbn(true);

      try {
        const url =
          `https://www.googleapis.com/books/v1/volumes?q=isbn:${encodeURIComponent(
            query
          )}&maxResults=10`;

        const response =
          await fetch(url);

        if (!response.ok) {
          throw new Error(
            `HTTP ${response.status}`
          );
        }

        const data =
          await response.json();

        if (
          !data.items ||
          data.items.length === 0
        ) {
          alert(
            `Nessun libro trovato per ISBN ${query}.`
          );
          return;
        }

        /*
         * Cerchiamo prima una corrispondenza
         * esatta dell'ISBN.
         */
        let selected =
          data.items[0];

        for (
          const item of data.items
        ) {
          const identifiers =
            item.volumeInfo
              ?.industryIdentifiers ||
            [];

          const exact =
            identifiers.some(
              (identifier: any) =>
                normalizeISBN(
                  identifier.identifier
                ) === query
            );

          if (exact) {
            selected = item;
            break;
          }
        }

        const info =
          selected.volumeInfo || {};

        const identifiers =
          info.industryIdentifiers ||
          [];

        const isbn13 =
          identifiers.find(
            (item: any) =>
              item.type === 'ISBN_13'
          )?.identifier;

        const isbn10 =
          identifiers.find(
            (item: any) =>
              item.type === 'ISBN_10'
          )?.identifier;

        const finalISBN =
          isbn13 ||
          isbn10 ||
          query;

        const image =
          info.imageLinks
            ?.thumbnail ||
          info.imageLinks
            ?.smallThumbnail ||
          '';

        setFormData(
          (previous) => ({
            ...previous,

            title:
              info.title ||
              previous.title ||
              '',

            author:
              info.authors?.join(
                ', '
              ) ||
              previous.author ||
              '',

            publisher:
              info.publisher ||
              previous.publisher ||
              '',

            publishYear:
              info.publishedDate
                ? info.publishedDate.substring(
                    0,
                    4
                  )
                : previous.publishYear ||
                  '',

            pages:
              info.pageCount ||
              previous.pages ||
              undefined,

            genre:
              info.categories?.[0] ||
              previous.genre ||
              '',

            coverUrl:
              image
                ? image.replace(
                    'http:',
                    'https:'
                  )
                : previous.coverUrl ||
                  '',

            isbn:
              finalISBN,
          })
        );

        setIsbnInput(
          finalISBN
        );
      } catch (error) {
        console.error(
          'Errore Google Books:',
          error
        );

        alert(
          'Errore durante la ricerca online. Controlla la connessione Internet.'
        );
      } finally {
        setIsSearchingIsbn(false);
      }
    };

  /* =======================================================
     RESET FORM
  ======================================================= */

  const resetForm = () => {
    setFormData({
      title: '',
      author: '',
      publishCountry: '',
      coverUrl: '',
      publisher: '',
      publishYear: '',
      pages: undefined,
      genre: '',
      seriesTag: '',
      volume: '',
      isClassic: false,
      format: 'cartaceo',
      isRead: false,
      readMonth: 'Gennaio',
      readYear: currentYearNum,
      rating: 5,
      notes: '',
      isbn: '',
    });

    setIsbnInput('');
  };

  /* =======================================================
     SALVA LIBRO
  ======================================================= */

  const handleSaveBook = (
    event: React.FormEvent
  ) => {
    event.preventDefault();

    if (
      !formData.title?.trim() ||
      !formData.author?.trim()
    ) {
      alert(
        'Inserisci almeno Titolo e Autore.'
      );

      return;
    }

    const formattedReadMonthYear =
      formData.isRead &&
      formData.readMonth &&
      formData.readYear
        ? `${formData.readMonth} ${formData.readYear}`
        : '';

    const newBook: BookItem = {
      id:
        formData.id ||
        Date.now().toString(),

      title:
        formData.title.trim(),

      author:
        formData.author.trim(),

      publishCountry:
        formData.publishCountry ||
        '',

      coverUrl:
        formData.coverUrl ||
        '',

      publisher:
        formData.publisher ||
        '',

      publishYear:
        formData.publishYear ||
        '',

      pages:
        Number(formData.pages) ||
        undefined,

      genre:
        formData.genre ||
        '',

      seriesTag:
        formData.seriesTag ||
        '',

      volume:
        formData.volume ||
        '',

      isClassic:
        Boolean(formData.isClassic),

      format:
        formData.format ||
        'cartaceo',

      isRead:
        Boolean(formData.isRead),

      readMonth:
        formData.readMonth ||
        'Gennaio',

      readYear:
        formData.readYear ||
        currentYearNum,

      readMonthYear:
        formattedReadMonthYear,

      rating:
        Number(formData.rating) ||
        5,

      notes:
        formData.notes ||
        '',

      isbn:
        formData.isbn ||
        isbnInput ||
        '',

      createdAt:
        formData.createdAt ||
        Date.now(),
    };

    if (formData.id) {
      setBooks(
        (previous) =>
          previous.map((book) =>
            book.id === formData.id
              ? newBook
              : book
          )
      );
    } else {
      setBooks(
        (previous) => [
          newBook,
          ...previous,
        ]
      );
    }

    setIsAddModalOpen(false);
    setSelectedBookDetail(null);

    resetForm();
  };

  /* =======================================================
     MODIFICA
  ======================================================= */

  const handleEditBook = (
    book: BookItem
  ) => {
    setFormData({
      ...book,
    });

    setIsbnInput(
      book.isbn || ''
    );

    setSelectedBookDetail(null);

    setIsAddModalOpen(true);
  };

  /* =======================================================
     ELIMINA
  ======================================================= */

  const handleDeleteBook = (
    id: string
  ) => {
    if (
      !confirm(
        'Sei sicuro di voler eliminare questo libro?'
      )
    ) {
      return;
    }

    setBooks(
      (previous) =>
        previous.filter(
          (book) =>
            book.id !== id
        )
    );

    if (
      selectedBookDetail?.id === id
    ) {
      setSelectedBookDetail(null);
    }
  };

  /* =======================================================
     DRAG & DROP
  ======================================================= */

  const handleDragStart = (
    index: number
  ) => {
    setDraggedIndex(index);
  };

  const handleDragOver = (
    event: React.DragEvent,
    index: number
  ) => {
    event.preventDefault();

    if (
      draggedIndex === null ||
      draggedIndex === index
    ) {
      return;
    }

    const readOnlyBooks =
      books.filter(
        (book) => book.isRead
      );

    const itemToMove =
      readOnlyBooks[draggedIndex];

    const target =
      readOnlyBooks[index];

    if (!itemToMove || !target) {
      return;
    }

    const updated =
      [...books];

    const sourceGlobalIndex =
      updated.findIndex(
        (book) =>
          book.id ===
          itemToMove.id
      );

    const targetGlobalIndex =
      updated.findIndex(
        (book) =>
          book.id ===
          target.id
      );

    if (
      sourceGlobalIndex < 0 ||
      targetGlobalIndex < 0
    ) {
      return;
    }

    updated.splice(
      sourceGlobalIndex,
      1
    );

    updated.splice(
      targetGlobalIndex,
      0,
      itemToMove
    );

    setBooks(updated);

    setDraggedIndex(index);
  };

  const handleDragEnd = () => {
    setDraggedIndex(null);
  };

  /* =======================================================
     EXPORT EXCEL
  ======================================================= */

  const exportToExcel = () => {
    const data =
      books.map((book) => ({
        Titolo: book.title,
        Autore: book.author,
        'Paese di Pubblicazione':
          book.publishCountry || '-',
        Classico:
          book.isClassic
            ? 'Sì'
            : 'No',
        Stato:
          book.isRead
            ? 'Letto'
            : 'In Biblioteca',
        Formato:
          book.format ===
          'cartaceo'
            ? 'Cartaceo'
            : book.format ===
              'ebook'
            ? 'eBook'
            : 'eBook + Cartaceo',
        Editore:
          book.publisher || '-',
        'Anno Pubblicazione':
          book.publishYear || '-',
        Genere:
          book.genre || '-',
        'Serie / Tag':
          book.seriesTag || '-',
        Volume:
          book.volume || '-',
        Pagine:
          book.pages || '-',
        'Mese e Anno di Lettura':
          book.readMonthYear ||
          '-',
        ISBN:
          book.isbn || '-',
      }));

    const worksheet =
      XLSX.utils.json_to_sheet(
        data
      );

    const workbook =
      XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Biblioteca'
    );

    XLSX.writeFile(
      workbook,
      'La_Mia_Biblioteca.xlsx'
    );
  };

  /* =======================================================
     EXPORT PDF
  ======================================================= */

  const exportToPDF = () => {
    const doc =
      new jsPDF();

    doc.text(
      'La Mia Biblioteca - Report',
      14,
      15
    );

    const tableData =
      books.map((book) => [
        book.title,
        book.author,
        book.publishCountry ||
          '-',
        book.isClassic
          ? 'Sì'
          : 'No',
        book.isRead
          ? 'Letto'
          : 'In Libreria',
        book.format ===
        'cartaceo'
          ? 'Cartaceo'
          : book.format ===
            'ebook'
          ? 'eBook'
          : 'eBook + Cartaceo',
        book.genre || '-',
        book.readMonthYear ||
          '-',
      ]);

    autoTable(doc, {
      head: [
        [
          'Titolo',
          'Autore',
          'Paese',
          'Classico',
          'Stato',
          'Formato',
          'Genere',
          'Data Lettura',
        ],
      ],
      body: tableData,
      startY: 20,
    });

    doc.save(
      'La_Mia_Biblioteca.pdf'
    );
  };

  /* =======================================================
     BACKUP
  ======================================================= */

  const exportBackup = () => {
    const data =
      'data:text/json;charset=utf-8,' +
      encodeURIComponent(
        JSON.stringify(
          books
        )
      );

    const anchor =
      document.createElement('a');

    anchor.href = data;

    anchor.download =
      'backup_libreria.json';

    document.body.appendChild(
      anchor
    );

    anchor.click();

    anchor.remove();
  };

  const handleImportBackup = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file =
      event.target.files?.[0];

    if (!file) return;

    const reader =
      new FileReader();

    reader.onload = () => {
      try {
        const parsed =
          JSON.parse(
            reader.result as string
          );

        if (
          !Array.isArray(parsed)
        ) {
          throw new Error(
            'Formato non valido'
          );
        }

        setBooks(parsed);

        alert(
          'Backup ripristinato con successo!'
        );
      } catch {
        alert(
          'File di backup non valido.'
        );
      }
    };

    reader.readAsText(
      file,
      'UTF-8'
    );
  };

  /* =======================================================
     CANCELLA TUTTO
  ======================================================= */

  const handleClearAll = () => {
    if (
      !confirm(
        'ATTENZIONE: Verranno cancellati TUTTI i libri salvati. Procedere?'
      )
    ) {
      return;
    }

    setBooks([]);

    localStorage.removeItem(
      STORAGE_KEY
    );
  };

  /* =======================================================
     STATISTICHE
  ======================================================= */

  const totalBooks =
    books.length;

  const totalCartacei =
    books.filter(
      (book) =>
        book.format ===
          'cartaceo' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const totalEbook =
    books.filter(
      (book) =>
        book.format ===
          'ebook' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readBooks =
    books.filter(
      (book) =>
        book.isRead
    );

  const readBooksCount =
    readBooks.length;

  const readCartacei =
    readBooks.filter(
      (book) =>
        book.format ===
          'cartaceo' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readEbook =
    readBooks.filter(
      (book) =>
        book.format ===
          'ebook' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readThisYearBooks =
    books.filter((book) => {
      if (!book.isRead) {
        return false;
      }

      if (book.readYear) {
        return (
          book.readYear ===
          currentYearNum
        );
      }

      return Boolean(
        book.readMonthYear?.includes(
          currentYearNum.toString()
        )
      );
    });

  const readThisYearCount =
    readThisYearBooks.length;

  const readThisYearCartacei =
    readThisYearBooks.filter(
      (book) =>
        book.format ===
          'cartaceo' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readThisYearEbook =
    readThisYearBooks.filter(
      (book) =>
        book.format ===
          'ebook' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  /* =======================================================
     FILTRI LIBRI LETTI
  ======================================================= */

  const readBooksFiltered =
    books
      .filter(
        (book) =>
          book.isRead
      )
      .filter(
        (book) =>
          filterGenre ===
          'all'
            ? true
            : book.genre ===
              filterGenre
      )
      .filter((book) => {
        if (
          filterFormat ===
          'all'
        ) {
          return true;
        }

        if (
          filterFormat ===
          'cartaceo'
        ) {
          return (
            book.format ===
            'cartaceo'
          );
        }

        if (
          filterFormat ===
          'ebook'
        ) {
          return (
            book.format ===
              'ebook' ||
            book.format ===
              'ebook_and_paper'
          );
        }

        return true;
      })
      .filter((book) => {
        if (
          filterYear ===
          'all'
        ) {
          return true;
        }

        return (
          book.readYear?.toString() ===
            filterYear ||
          book.readMonthYear?.includes(
            filterYear
          )
        );
      })
      .filter((book) => {
        if (
          searchQuery ===
          ''
        ) {
          return true;
        }

        const query =
          searchQuery.toLowerCase();

        return (
          book.title
            .toLowerCase()
            .includes(query) ||
          book.author
            .toLowerCase()
            .includes(query)
        );
      });

  /* =======================================================
     AUTORI
  ======================================================= */

  const authorsMap =
    books.reduce(
      (
        accumulator,
        book
      ) => {
        const name =
          book.author.trim() ||
          'Autore Sconosciuto';

        if (!accumulator[name]) {
          accumulator[name] =
            [];
        }

        accumulator[name].push(
          book
        );

        return accumulator;
      },
      {} as Record<
        string,
        BookItem[]
      >
    );

  const sortedAuthors =
    Object.keys(
      authorsMap
    ).sort((a, b) =>
      a.localeCompare(b)
    );

  /* =======================================================
     GENERI
  ======================================================= */

  const genresMap =
    books.reduce(
      (
        accumulator,
        book
      ) => {
        const name =
          book.genre?.trim() ||
          'Generico / Altro';

        if (!accumulator[name]) {
          accumulator[name] =
            [];
        }

        accumulator[name].push(
          book
        );

        return accumulator;
      },
      {} as Record<
        string,
        BookItem[]
      >
    );

  const sortedGenres =
    Object.keys(
      genresMap
    ).sort((a, b) =>
      a.localeCompare(b)
    );

  const availableGenres =
    Array.from(
      new Set(
        books
          .map(
            (book) =>
              book.genre
          )
          .filter(Boolean)
      )
    );

  const availableYears =
    Array.from(
      new Set(
        books
          .map((book) => {
            if (
              book.readYear
            ) {
              return book.readYear.toString();
            }

            const match =
              book.readMonthYear?.match(
                /\d{4}/
              );

            return match
              ? match[0]
              : null;
          })
          .filter(Boolean)
      )
    ) as string[];

  /* =======================================================
     FORMAT LABEL
  ======================================================= */

  const formatLabel = (
    format: string
  ) => {
    if (
      format ===
      'cartaceo'
    ) {
      return 'Cartaceo';
    }

    if (
      format ===
      'ebook'
    ) {
      return 'eBook';
    }

    if (
      format ===
      'ebook_and_paper'
    ) {
      return 'eBook + Cartaceo';
    }

    return format;
  };

  /* =======================================================
     GRUPPI AUTORI
  ======================================================= */

  const renderAuthorGroup =
    (
      bookList: BookItem[]
    ) => {
      const map =
        bookList.reduce(
          (
            accumulator,
            book
          ) => {
            const author =
              book.author.trim() ||
              'Autore Sconosciuto';

            if (
              !accumulator[
                author
              ]
            ) {
              accumulator[
                author
              ] = [];
            }

            accumulator[
              author
            ].push(book);

            return accumulator;
          },
          {} as Record<
            string,
            BookItem[]
          >
        );

      return Object.keys(
        map
      )
        .sort()
        .map(
          (author) => (
            <div
              key={author}
              className="space-y-2 pt-2"
            >
              <h3 className="font-serif font-bold text-sm text-amber-950 border-b border-amber-900/10 pb-1">
                {author}
              </h3>

              <div className="grid grid-cols-2 gap-3">
                {map[author]
                  .sort(
                    (a, b) =>
                      (
                        a.volume ||
                        ''
                      ).localeCompare(
                        b.volume ||
                          ''
                      )
                  )
                  .map(
                    (book) => (
                      <div
                        key={
                          book.id
                        }
                        onClick={() =>
                          setSelectedBookDetail(
                            book
                          )
                        }
                        className="bg-[#FFFDF9] p-3 rounded-2xl shadow-sm border border-amber-900/10 flex flex-col cursor-pointer"
                      >
                        <div className="w-full h-36 bg-amber-100/40 rounded-xl overflow-hidden relative mb-2 border border-amber-900/10">
                          {book.coverUrl ? (
                            <img
                              src={
                                book.coverUrl
                              }
                              alt={
                                book.title
                              }
                              className="w-full h-full object-cover"
                            />
                          ) : (
                            <div className="w-full h-full flex items-center justify-center text-amber-800/30">
                              <Book className="w-6 h-6" />
                            </div>
                          )}

                          {book.isRead && (
                            <div className="absolute top-2 right-2 bg-emerald-700 text-amber-50 p-1 rounded-full shadow">
                              <CheckCircle2 className="w-3 h-3" />
                            </div>
                          )}

                          {book.volume && (
                            <div className="absolute top-2 left-2 bg-amber-800 text-amber-50 text-[9px] font-bold px-2 py-0.5 rounded-full">
                              Vol.{' '}
                              {
                                book.volume
                              }
                            </div>
                          )}
                        </div>

                        <h4 className="font-serif font-bold text-xs text-amber-950 line-clamp-2">
                          {
                            book.title
                          }
                        </h4>

                        <p className="text-[10px] text-amber-800/60 mt-0.5">
                          {book.publishYear
                            ? `Anno: ${book.publishYear}`
                            : ''}
                        </p>
                      </div>
                    )
                  )}
              </div>
            </div>
          )
        );
    };

  /* =======================================================
     RENDER
  ======================================================= */

  return (
    <div className="min-h-screen bg-[#FBF9F5] text-amber-950 font-sans pb-24 select-none">

      {/* =================================================
          HEADER
      ================================================= */}

      <header className="sticky top-0 z-20 bg-[#FBF9F5]/90 backdrop-blur-md border-b border-amber-900/10 px-5 py-3.5 flex justify-between items-center">
        <div>
          <span className="text-[11px] font-bold text-amber-800/70 uppercase tracking-widest block">
            {activeTab ===
              'home' &&
              'La Mia Collezione'}

            {activeTab ===
              'read' &&
              'Cronologia Letture'}

            {activeTab ===
              'authors' &&
              'Catalogo Autori'}

            {activeTab ===
              'settings' &&
              'Gestione Dati'}
          </span>

          <h1 className="text-2xl font-serif font-extrabold tracking-tight text-amber-950">
            {activeTab ===
              'home' &&
              'Home'}

            {activeTab ===
              'read' &&
              'Libri Letti'}

            {activeTab ===
              'authors' &&
              (selectedAuthor ||
                'Autori')}

            {activeTab ===
              'settings' &&
              'Impostazioni'}
          </h1>
        </div>

        {activeTab ===
          'read' && (
          <button
            onClick={() => {
              resetForm();
              setIsAddModalOpen(
                true
              );
            }}
            className="w-10 h-10 bg-amber-800 text-amber-50 rounded-full flex items-center justify-center shadow-md active:scale-95 transition-transform"
          >
            <Plus className="w-5 h-5 stroke-[2.5]" />
          </button>
        )}
      </header>

      {/* =================================================
          HOME
      ================================================= */}

      {activeTab ===
        'home' && (
        <div className="p-4 space-y-5 max-w-lg mx-auto">
          {homeSubView !==
          'none' ? (
            <div className="space-y-4">
              <button
                onClick={() => {
                  setHomeSubView(
                    'none'
                  );
                  setSelectedGenreHome(
                    null
                  );
                }}
                className="text-xs font-bold text-amber-800"
              >
                ← Torna alla Home
              </button>

              {homeSubView ===
                'classics' && (
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <Sparkles className="w-5 h-5 text-amber-600" />
                    <h2 className="text-lg font-serif font-bold">
                      I Miei Classici
                    </h2>
                  </div>

                  {renderAuthorGroup(
                    books.filter(
                      (book) =>
                        book.isClassic
                    )
                  )}
                </div>
              )}

              {homeSubView ===
                'genres' && (
                <div className="space-y-4">
                  {!selectedGenreHome ? (
                    <div className="space-y-2">
                      <h2 className="text-lg font-serif font-bold mb-2">
                        Generi nella Biblioteca
                      </h2>

                      {sortedGenres.map(
                        (
                          genre
                        ) => (
                          <div
                            key={
                              genre
                            }
                            onClick={() =>
                              setSelectedGenreHome(
                                genre
                              )
                            }
                            className="bg-[#FFFDF9] p-4 rounded-2xl shadow-sm border border-amber-900/10 flex justify-between items-center cursor-pointer"
                          >
                            <div>
                              <h3 className="font-serif font-bold">
                                {
                                  genre
                                }
                              </h3>

                              <p className="text-xs text-amber-800/60">
                                {
                                  genresMap[
                                    genre
                                  ]
                                    .length
                                }{' '}
                                libri
                              </p>
                            </div>

                            <ChevronRight className="w-5 h-5 text-amber-800/30" />
                          </div>
                        )
                      )}
                    </div>
                  ) : (
                    <div className="space-y-3">
                      <button
                        onClick={() =>
                          setSelectedGenreHome(
                            null
                          )
                        }
                        className="text-xs font-bold text-amber-800"
                      >
                        ← Tutti i generi
                      </button>

                      <h2 className="text-lg font-serif font-bold">
                        {
                          selectedGenreHome
                        }
                      </h2>

                      {renderAuthorGroup(
                        genresMap[
                          selectedGenreHome
                        ] || []
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>
          ) : (
            <>
              <div className="bg-[#FFFDF9] rounded-3xl p-5 shadow-sm border border-amber-900/10 space-y-3">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 rounded-2xl bg-amber-800 text-amber-50 flex items-center justify-center shadow-md">
                    <Calendar className="w-6 h-6" />
                  </div>

                  <div>
                    <span className="text-xs font-bold text-amber-800/60 uppercase tracking-wider block">
                      Anno{' '}
                      {
                        currentYearNum
                      }
                    </span>

                    <span className="text-2xl font-serif font-bold">
                      {
                        readThisYearCount
                      }{' '}
                      <span className="text-sm font-sans font-normal text-amber-800/70">
                        libri letti
                      </span>
                    </span>
                  </div>
                </div>

                <div className="pt-2.5 border-t border-amber-900/10 flex justify-center text-xs font-semibold">
                  {
                    readThisYearCartacei
                  }{' '}
                  Cartacei e{' '}
                  {
                    readThisYearEbook
                  }{' '}
                  eBook
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3.5">
                <div className="bg-[#FFFDF9] p-4 rounded-3xl shadow-sm border border-amber-900/10 min-h-[9rem]">
                  <BookOpen className="w-5 h-5" />

                  <div className="mt-5">
                    <span className="text-3xl font-serif font-black">
                      {
                        totalBooks
                      }
                    </span>

                    <span className="text-xs font-bold uppercase tracking-wider block mt-1">
                      In Biblioteca
                    </span>

                    <div className="mt-2 pt-2 border-t border-amber-900/10 text-xs font-semibold">
                      {
                        totalCartacei
                      }{' '}
                      Cartacei e{' '}
                      {
                        totalEbook
                      }{' '}
                      eBook
                    </div>
                  </div>
                </div>

                <div className="bg-[#FFFDF9] p-4 rounded-3xl shadow-sm border border-amber-900/10 min-h-[9rem]">
                  <CheckCircle2 className="w-5 h-5 text-emerald-700" />

                  <div className="mt-5">
                    <span className="text-3xl font-serif font-black">
                      {
                        readBooksCount
                      }
                    </span>

                    <span className="text-xs font-bold uppercase tracking-wider block mt-1">
                      Libri Letti
                    </span>

                    <div className="mt-2 pt-2 border-t border-amber-900/10 text-xs font-semibold">
                      {
                        readCartacei
                      }{' '}
                      Cartacei e{' '}
                      {
                        readEbook
                      }{' '}
                      eBook
                    </div>
                  </div>
                </div>
              </div>

              <div className="space-y-3 pt-2">
                <h2 className="text-xs font-bold text-amber-800/60 uppercase tracking-wider px-1">
                  Esplora Categorie
                </h2>

                <div className="grid grid-cols-2 gap-3.5">
                  <button
                    onClick={() =>
                      setHomeSubView(
                        'classics'
                      )
                    }
                    className="text-left bg-gradient-to-br from-amber-700 to-amber-900 text-amber-50 p-4 rounded-3xl shadow-md h-32 flex flex-col justify-between"
                  >
                    <Sparkles className="w-6 h-6" />

                    <div>
                      <span className="text-lg font-serif font-bold block">
                        Classici
                      </span>

                      <span className="text-xs text-amber-200/80">
                        {
                          books.filter(
                            (book) =>
                              book.isClassic
                          ).length
                        }{' '}
                        libri
                      </span>
                    </div>
                  </button>

                  <button
                    onClick={() =>
                      setHomeSubView(
                        'genres'
                      )
                    }
                    className="text-left bg-gradient-to-br from-stone-800 to-amber-950 text-amber-50 p-4 rounded-3xl shadow-md h-32 flex flex-col justify-between"
                  >
                    <Layers className="w-6 h-6" />

                    <div>
                      <span className="text-lg font-serif font-bold block">
                        Generi
                      </span>

                      <span className="text-xs text-amber-200/80">
                        {
                          sortedGenres.length
                        }{' '}
                        categorie
                      </span>
                    </div>
                  </button>
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* =================================================
          LIBRI LETTI
      ================================================= */}

      {activeTab ===
        'read' && (
        <div className="p-4 space-y-4 max-w-lg mx-auto">
          <div className="relative">
            <Search className="w-4 h-4 absolute left-3.5 top-3 text-amber-800/40" />

            <input
              value={searchQuery}
              onChange={(e) =>
                setSearchQuery(
                  e.target.value
                )
              }
              placeholder="Cerca nei libri letti..."
              className="w-full pl-9 pr-4 py-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-2xl text-xs focus:outline-none"
            />
          </div>

          <div className="flex gap-2 overflow-x-auto pb-1">
            <select
              value={
                filterGenre
              }
              onChange={(e) =>
                setFilterGenre(
                  e.target.value
                )
              }
              className="bg-[#FFFDF9] border border-amber-900/10 rounded-xl px-3 py-2 text-xs"
            >
              <option value="all">
                Tutti i generi
              </option>

              {availableGenres.map(
                (genre) => (
                  <option
                    key={genre}
                    value={genre}
                  >
                    {genre}
                  </option>
                )
              )}
            </select>

            <select
              value={
                filterFormat
              }
              onChange={(e) =>
                setFilterFormat(
                  e.target.value
                )
              }
              className="bg-[#FFFDF9] border border-amber-900/10 rounded-xl px-3 py-2 text-xs"
            >
              <option value="all">
                Tutti i formati
              </option>

              <option value="cartaceo">
                Cartaceo
              </option>

              <option value="ebook">
                eBook
              </option>
            </select>

            <select
              value={
                filterYear
              }
              onChange={(e) =>
                setFilterYear(
                  e.target.value
                )
              }
              className="bg-[#FFFDF9] border border-amber-900/10 rounded-xl px-3 py-2 text-xs"
            >
              <option value="all">
                Tutti gli anni
              </option>

              {availableYears.map(
                (year) => (
                  <option
                    key={year}
                    value={year}
                  >
                    {year}
                  </option>
                )
              )}
            </select>
          </div>

          <p className="text-[11px] text-amber-800/50 px-1">
            Trascina per riordinare i
            libri letti.
          </p>

          <div className="space-y-3">
            {readBooksFiltered.length ===
            0 ? (
              <div className="text-center py-12 bg-[#FFFDF9] rounded-3xl border border-dashed border-amber-900/20">
                <BookOpen className="w-8 h-8 text-amber-800/30 mx-auto mb-2" />

                <p className="text-xs text-amber-800/60">
                  Nessun libro
                  trovato.
                </p>
              </div>
            ) : (
              readBooksFiltered.map(
                (
                  book,
                  index
                ) => (
                  <div
                    key={
                      book.id
                    }
                    draggable
                    onDragStart={() =>
                      handleDragStart(
                        index
                      )
                    }
                    onDragOver={(
                      e
                    ) =>
                      handleDragOver(
                        e,
                        index
                      )
                    }
                    onDragEnd={
                      handleDragEnd
                    }
                    className={`bg-[#FFFDF9] p-3 rounded-2xl shadow-sm border border-amber-900/10 flex items-center gap-3 ${
                      draggedIndex ===
                      index
                        ? 'bg-amber-100/50'
                        : ''
                    }`}
                  >
                    <GripVertical className="w-5 h-5 text-amber-800/30 cursor-grab" />

                    <div
                      onClick={() =>
                        setSelectedBookDetail(
                          book
                        )
                      }
                      className="w-13 h-19 bg-amber-100/40 rounded-lg overflow-hidden flex-shrink-0"
                    >
                      {book.coverUrl ? (
                        <img
                          src={
                            book.coverUrl
                          }
                          alt={
                            book.title
                          }
                          className="w-full h-full object-cover"
                        />
                      ) : (
                        <div className="w-full h-full flex items-center justify-center">
                          <Book className="w-5 h-5 text-amber-800/30" />
                        </div>
                      )}
                    </div>

                    <div
                      onClick={() =>
                        setSelectedBookDetail(
                          book
                        )
                      }
                      className="flex-1 min-w-0 cursor-pointer"
                    >
                      <h3 className="font-serif font-bold text-sm truncate">
                        {
                          book.title
                        }
                      </h3>

                      <p className="text-xs text-amber-800/70 truncate">
                        {
                          book.author
                        }
                      </p>

                      {book.rating && (
                        <div className="flex text-amber-500 mt-1">
                          {Array.from(
                            {
                              length: 5,
                            }
                          ).map(
                            (
                              _,
                              i
                            ) => (
                              <Star
                                key={
                                  i
                                }
                                className={`w-3 h-3 ${
                                  i <
                                  book.rating!
                                    ? 'fill-amber-500'
                                    : 'text-amber-200'
                                }`}
                              />
                            )
                          )}
                        </div>
                      )}

                      <p className="text-[10px] text-amber-800/50 mt-1">
                        Letto:{' '}
                        <span className="text-amber-950">
                          {book.readMonthYear ||
                            'Data non specificata'}
                        </span>
                      </p>
                    </div>

                    <ChevronRight className="w-4 h-4 text-amber-800/30" />
                  </div>
                )
              )
            )}
          </div>
        </div>
      )}

      {/* =================================================
          AUTORI
      ================================================= */}

      {activeTab ===
        'authors' && (
        <div className="p-4 space-y-4 max-w-lg mx-auto">
          {!selectedAuthor ? (
            <div className="space-y-2">
              {sortedAuthors.map(
                (author) => (
                  <button
                    key={author}
                    onClick={() =>
                      setSelectedAuthor(
                        author
                      )
                    }
                    className="w-full text-left bg-[#FFFDF9] p-4 rounded-2xl shadow-sm border border-amber-900/10 flex justify-between items-center"
                  >
                    <div>
                      <h3 className="font-serif font-bold">
                        {author}
                      </h3>

                      <p className="text-xs text-amber-800/60">
                        {
                          authorsMap[
                            author
                          ].length
                        }{' '}
                        libri
                      </p>
                    </div>

                    <ChevronRight className="w-5 h-5 text-amber-800/30" />
                  </button>
                )
              )}
            </div>
          ) : (
            <div className="space-y-4">
              <button
                onClick={() =>
                  setSelectedAuthor(
                    null
                  )
                }
                className="text-xs font-bold text-amber-800"
              >
                ← Torna agli autori
              </button>

              <div className="grid grid-cols-2 gap-3">
                {authorsMap[
                  selectedAuthor
                ]?.map(
                  (book) => (
                    <button
                      key={
                        book.id
                      }
                      onClick={() =>
                        setSelectedBookDetail(
                          book
                        )
                      }
                      className="text-left bg-[#FFFDF9] p-3 rounded-2xl shadow-sm border border-amber-900/10"
                    >
                      <div className="w-full h-44 bg-amber-100/40 rounded-xl overflow-hidden mb-2">
                        {book.coverUrl ? (
                          <img
                            src={
                              book.coverUrl
                            }
                            alt={
                              book.title
                            }
                            className="w-full h-full object-cover"
                          />
                        ) : (
                          <div className="w-full h-full flex items-center justify-center">
                            <Book className="w-8 h-8 text-amber-800/30" />
                          </div>
                        )}
                      </div>

                      <h4 className="font-serif font-bold text-xs line-clamp-2">
                        {
                          book.title
                        }
                      </h4>

                      <p className="text-[10px] text-amber-800/60 mt-1">
                        {book.publishYear
                          ? `Anno: ${book.publishYear}`
                          : ''}
                      </p>
                    </button>
                  )
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {/* =================================================
          SETTINGS
      ================================================= */}

      {activeTab ===
        'settings' && (
        <div className="p-4 space-y-5 max-w-lg mx-auto">
          <div className="bg-[#FFFDF9] rounded-3xl p-4 shadow-sm border border-amber-900/10 space-y-3">
            <h2 className="text-xs font-bold text-amber-800/60 uppercase">
              Esporta
            </h2>

            <button
              onClick={
                exportToExcel
              }
              className="w-full p-3 bg-emerald-50 text-emerald-900 rounded-2xl font-bold text-xs flex items-center gap-3"
            >
              <FileSpreadsheet className="w-5 h-5" />
              Esporta Excel
            </button>

            <button
              onClick={
                exportToPDF
              }
              className="w-full p-3 bg-rose-50 text-rose-900 rounded-2xl font-bold text-xs flex items-center gap-3"
            >
              <FileText className="w-5 h-5" />
              Esporta PDF
            </button>
          </div>

          <div className="bg-[#FFFDF9] rounded-3xl p-4 shadow-sm border border-amber-900/10 space-y-3">
            <h2 className="text-xs font-bold text-amber-800/60 uppercase">
              Backup
            </h2>

            <button
              onClick={
                exportBackup
              }
              className="w-full p-3 bg-amber-100/60 rounded-2xl font-bold text-xs flex items-center gap-3"
            >
              <Download className="w-5 h-5" />
              Salva Backup JSON
            </button>

            <label className="w-full p-3 bg-stone-100 rounded-2xl font-bold text-xs flex items-center gap-3 cursor-pointer">
              <Upload className="w-5 h-5" />
              Ripristina Backup

              <input
                type="file"
                accept=".json"
                onChange={
                  handleImportBackup
                }
                className="hidden"
              />
            </label>
          </div>

          <div className="bg-[#FFFDF9] rounded-3xl p-4 shadow-sm border border-amber-900/10">
            <button
              onClick={
                handleClearAll
              }
              className="w-full p-3 bg-rose-700 text-white rounded-2xl font-bold text-xs flex items-center justify-center gap-2"
            >
              <Trash2 className="w-4 h-4" />
              Cancella Intera Biblioteca
            </button>
          </div>
        </div>
      )}

      {/* =================================================
          DETTAGLIO LIBRO
      ================================================= */}

      {selectedBookDetail && (
        <div className="fixed inset-0 z-50 bg-amber-950/40 backdrop-blur-sm flex items-end sm:items-center justify-center">
          <div className="bg-[#FFFDF9] w-full max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[90vh] overflow-y-auto p-6 space-y-4">
            <button
              onClick={() =>
                setSelectedBookDetail(
                  null
                )
              }
              className="absolute top-4 right-4 p-2 bg-amber-100 rounded-full"
            >
              <X className="w-5 h-5" />
            </button>

            <div className="flex gap-4 pt-2">
              <div className="w-24 h-36 bg-amber-100 rounded-xl overflow-hidden flex-shrink-0">
                {selectedBookDetail.coverUrl ? (
                  <img
                    src={
                      selectedBookDetail.coverUrl
                    }
                    alt={
                      selectedBookDetail.title
                    }
                    className="w-full h-full object-cover"
                  />
                ) : (
                  <div className="w-full h-full flex items-center justify-center">
                    <Book />
                  </div>
                )}
              </div>

              <div>
                <h2 className="text-lg font-serif font-bold">
                  {
                    selectedBookDetail.title
                  }
                </h2>

                <p className="text-xs text-amber-800/70">
                  {
                    selectedBookDetail.author
                  }
                </p>

                <div className="flex flex-wrap gap-2 mt-3">
                  <span className="text-[10px] px-2.5 py-1 rounded-full bg-amber-100 font-bold">
                    {formatLabel(
                      selectedBookDetail.format
                    )}
                  </span>

                  {selectedBookDetail.isClassic && (
                    <span className="text-[10px] px-2.5 py-1 rounded-full bg-amber-200 font-bold">
                      Classico
                    </span>
                  )}

                  <span className="text-[10px] px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-900 font-bold">
                    {selectedBookDetail.isRead
                      ? 'Letto'
                      : 'In Biblioteca'}
                  </span>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3 bg-amber-50 p-4 rounded-2xl text-xs">
              <div>
                <small>Paese</small>
                <strong className="block">
                  {selectedBookDetail.publishCountry ||
                    '-'}
                </strong>
              </div>

              <div>
                <small>Editore</small>
                <strong className="block">
                  {selectedBookDetail.publisher ||
                    '-'}
                </strong>
              </div>

              <div>
                <small>Anno</small>
                <strong className="block">
                  {selectedBookDetail.publishYear ||
                    '-'}
                </strong>
              </div>

              <div>
                <small>Pagine</small>
                <strong className="block">
                  {selectedBookDetail.pages ||
                    '-'}
                </strong>
              </div>

              <div>
                <small>Genere</small>
                <strong className="block">
                  {selectedBookDetail.genre ||
                    '-'}
                </strong>
              </div>

              <div>
                <small>ISBN</small>
                <strong className="block">
                  {selectedBookDetail.isbn ||
                    '-'}
                </strong>
              </div>

              <div>
                <small>Serie</small>
                <strong className="block">
                  {selectedBookDetail.seriesTag ||
                    '-'}
                </strong>
              </div>

              <div>
                <small>Volume</small>
                <strong className="block">
                  {selectedBookDetail.volume ||
                    '-'}
                </strong>
              </div>

              {selectedBookDetail.isRead && (
                <div className="col-span-2 border-t pt-2">
                  <small>
                    Data lettura
                  </small>

                  <strong className="block">
                    {
                      selectedBookDetail.readMonthYear
                    }
                  </strong>
                </div>
              )}
            </div>

            <div className="grid grid-cols-2 gap-3">
              <button
                onClick={() =>
                  handleEditBook(
                    selectedBookDetail
                  )
                }
                className="py-3 bg-amber-800 text-white rounded-xl font-bold text-xs flex items-center justify-center gap-2"
              >
                <Edit className="w-4 h-4" />
                Modifica
              </button>

              <button
                onClick={() =>
                  handleDeleteBook(
                    selectedBookDetail.id
                  )
                }
                className="py-3 bg-rose-100 text-rose-800 rounded-xl font-bold text-xs flex items-center justify-center gap-2"
              >
                <Trash2 className="w-4 h-4" />
                Elimina
              </button>
            </div>
          </div>
        </div>
      )}

      {/* =================================================
          FORM LIBRO
      ================================================= */}

      {isAddModalOpen && (
        <div className="fixed inset-0 z-50 bg-amber-950/40 backdrop-blur-sm flex items-end sm:items-center justify-center">
          <div className="bg-[#FFFDF9] w-full max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[92vh] overflow-y-auto p-6 space-y-4">

            <div className="flex justify-between items-center border-b pb-3">
              <h2 className="text-base font-serif font-bold">
                {formData.id
                  ? 'Modifica Libro'
                  : 'Aggiungi Nuovo Libro'}
              </h2>

              <button
                onClick={() =>
                  setIsAddModalOpen(
                    false
                  )
                }
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* ISBN */}
            <div className="bg-amber-100/50 p-3.5 rounded-2xl space-y-3 border">
              <div className="flex justify-between items-center">
                <span className="text-[10px] font-bold uppercase">
                  Compilazione Automatica ISBN
                </span>

                <Scan className="w-4 h-4" />
              </div>

              <div className="flex gap-2">
                <input
                  type="text"
                  inputMode="numeric"
                  placeholder="ISBN"
                  value={
                    isbnInput
                  }
                  onChange={(
                    event
                  ) =>
                    setIsbnInput(
                      event.target.value
                    )
                  }
                  className="flex-1 min-w-0 p-2.5 bg-white border rounded-xl text-xs"
                />

                <button
                  type="button"
                  onClick={
                    openScanner
                  }
                  className="w-12 flex-shrink-0 bg-amber-950 text-white rounded-xl flex items-center justify-center"
                  aria-label="Scansiona ISBN"
                >
                  <Scan className="w-5 h-5" />
                </button>

                <button
                  type="button"
                  onClick={() =>
                    handleSearchBookByISBN()
                  }
                  disabled={
                    isSearchingIsbn
                  }
                  className="px-4 bg-amber-800 text-white rounded-xl text-xs font-bold disabled:opacity-50"
                >
                  {isSearchingIsbn
                    ? '...'
                    : 'Cerca'}
                </button>
              </div>

              <p className="text-[10px] text-amber-900/60">
                Scansiona il codice
                EAN-13 sul retro
                del libro oppure
                inserisci l'ISBN
                manualmente.
              </p>
            </div>

            <form
              onSubmit={
                handleSaveBook
              }
              className="space-y-3 text-xs"
            >
              <input
                required
                placeholder="Titolo *"
                value={
                  formData.title ||
                  ''
                }
                onChange={(e) =>
                  setFormData({
                    ...formData,
                    title:
                      e.target.value,
                  })
                }
                className="w-full p-3 bg-white border rounded-xl"
              />

              <div className="grid grid-cols-2 gap-2">
                <input
                  required
                  placeholder="Autore *"
                  value={
                    formData.author ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      author:
                        e.target.value,
                    })
                  }
                  className="p-3 bg-white border rounded-xl"
                />

                <input
                  placeholder="Paese pubblicazione"
                  value={
                    formData.publishCountry ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      publishCountry:
                        e.target.value,
                    })
                  }
                  className="p-3 bg-white border rounded-xl"
                />
              </div>

              <input
                placeholder="URL Copertina"
                value={
                  formData.coverUrl ||
                  ''
                }
                onChange={(e) =>
                  setFormData({
                    ...formData,
                    coverUrl:
                      e.target.value,
                  })
                }
                className="w-full p-3 bg-white border rounded-xl"
              />

              <div className="grid grid-cols-2 gap-2">
                <input
                  placeholder="Editore"
                  value={
                    formData.publisher ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      publisher:
                        e.target.value,
                    })
                  }
                  className="p-3 bg-white border rounded-xl"
                />

                <input
                  placeholder="Anno Pubblicazione"
                  value={
                    formData.publishYear ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      publishYear:
                        e.target.value,
                    })
                  }
                  className="p-3 bg-white border rounded-xl"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <input
                  placeholder="Genere"
                  value={
                    formData.genre ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      genre:
                        e.target.value,
                    })
                  }
                  className="p-3 bg-white border rounded-xl"
                />

                <input
                  type="number"
                  min="1"
                  placeholder="Numero Pagine"
                  value={
                    formData.pages ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      pages:
                        e.target.value
                          ? Number(
                              e.target.value
                            )
                          : undefined,
                    })
                  }
                  className="p-3 bg-white border rounded-xl"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <input
                  placeholder="Serie / Tag"
                  value={
                    formData.seriesTag ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      seriesTag:
                        e.target.value,
                    })
                  }
                  className="p-3 bg-white border rounded-xl"
                />

                <input
                  placeholder="Volume"
                  value={
                    formData.volume ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      volume:
                        e.target.value,
                    })
                  }
                  className="p-3 bg-white border rounded-xl"
                />
              </div>

              <div className="flex items-center justify-between p-3 bg-amber-50 rounded-xl border">
                <span className="font-bold flex items-center gap-1.5">
                  <Sparkles className="w-4 h-4" />
                  È un Classico?
                </span>

                <input
                  type="checkbox"
                  checked={
                    Boolean(
                      formData.isClassic
                    )
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      isClassic:
                        e.target.checked,
                    })
                  }
                  className="w-5 h-5 accent-amber-800"
                />
              </div>

              <div className="p-3 bg-amber-50 rounded-xl space-y-2 border">
                <span className="font-bold">
                  Formato
                </span>

                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() =>
                      setFormData({
                        ...formData,
                        format:
                          'cartaceo',
                      })
                    }
                    className={`flex-1 py-2 rounded-lg font-bold ${
                      formData.format ===
                      'cartaceo'
                        ? 'bg-amber-800 text-white'
                        : 'bg-white border'
                    }`}
                  >
                    Cartaceo
                  </button>

                  <button
                    type="button"
                    onClick={() =>
                      setFormData({
                        ...formData,
                        format:
                          'ebook',
                      })
                    }
                    className={`flex-1 py-2 rounded-lg font-bold ${
                      formData.format ===
                        'ebook' ||
                      formData.format ===
                        'ebook_and_paper'
                        ? 'bg-amber-800 text-white'
                        : 'bg-white border'
                    }`}
                  >
                    eBook
                  </button>
                </div>

                {(formData.format ===
                  'ebook' ||
                  formData.format ===
                    'ebook_and_paper') && (
                  <label className="flex items-center gap-2 pt-1">
                    <input
                      type="checkbox"
                      checked={
                        formData.format ===
                        'ebook_and_paper'
                      }
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          format:
                            e.target
                              .checked
                              ? 'ebook_and_paper'
                              : 'ebook',
                        })
                      }
                    />

                    Acquistato anche
                    in formato
                    Cartaceo
                  </label>
                )}
              </div>

              <div className="flex items-center justify-between p-3 bg-amber-50 rounded-xl border">
                <span className="font-bold">
                  Letto
                </span>

                <input
                  type="checkbox"
                  checked={
                    Boolean(
                      formData.isRead
                    )
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      isRead:
                        e.target.checked,
                    })
                  }
                  className="w-5 h-5 accent-amber-800"
                />
              </div>

              {formData.isRead && (
                <div className="p-3 bg-amber-100/40 border rounded-xl space-y-2">
                  <span className="font-bold">
                    Data di Lettura
                  </span>

                  <div className="grid grid-cols-2 gap-2">
                    <select
                      value={
                        formData.readMonth ||
                        'Gennaio'
                      }
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          readMonth:
                            e.target
                              .value,
                        })
                      }
                      className="p-2.5 bg-white border rounded-xl"
                    >
                      {MONTHS.map(
                        (month) => (
                          <option
                            key={
                              month
                            }
                            value={
                              month
                            }
                          >
                            {
                              month
                            }
                          </option>
                        )
                      )}
                    </select>

                    <select
                      value={
                        formData.readYear ||
                        currentYearNum
                      }
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          readYear:
                            Number(
                              e.target
                                .value
                            ),
                        })
                      }
                      className="p-2.5 bg-white border rounded-xl"
                    >
                      {yearsList.map(
                        (year) => (
                          <option
                            key={
                              year
                            }
                            value={
                              year
                            }
                          >
                            {
                              year
                            }
                          </option>
                        )
                      )}
                    </select>
                  </div>
                </div>
              )}

              <button
                type="submit"
                className="w-full py-3 bg-amber-800 text-white rounded-xl font-bold shadow-md"
              >
                {formData.id
                  ? 'Aggiorna Libro'
                  : 'Salva Libro'}
              </button>
            </form>
          </div>
        </div>
      )}

      {/* =================================================
          MODALE SCANNER
      ================================================= */}

      {isScannerOpen && (
        <div className="fixed inset-0 z-[100] bg-black flex flex-col">

          <div className="relative flex-1 overflow-hidden bg-black">

            <video
              ref={scannerVideoRef}
              autoPlay
              muted
              playsInline
              className="absolute inset-0 w-full h-full object-cover"
            />

            {/* Maschera */}
            <div className="absolute inset-0 pointer-events-none">

              <div className="absolute inset-0 bg-black/30" />

              <div className="absolute left-6 right-6 top-1/2 -translate-y-1/2 h-[180px]">

                <div className="absolute inset-0 border-2 border-white rounded-2xl shadow-[0_0_0_9999px_rgba(0,0,0,0.35)]" />

                <div className="absolute left-3 right-3 top-1/2 h-0.5 bg-red-500 shadow-lg" />

              </div>
            </div>

            {/* HEADER SCANNER */}
            <div className="absolute top-0 left-0 right-0 pt-[calc(env(safe-area-inset-top)+16px)] px-5">
              <div className="flex items-center justify-between">

                <button
                  type="button"
                  onClick={
                    closeScanner
                  }
                  className="w-11 h-11 rounded-full bg-black/60 text-white flex items-center justify-center backdrop-blur-md"
                >
                  <X className="w-5 h-5" />
                </button>

                <span className="text-white font-bold text-sm bg-black/60 px-4 py-2 rounded-full backdrop-blur-md">
                  Scansione ISBN
                </span>

                <div className="w-11" />
              </div>
            </div>

            {/* STATO */}
            <div className="absolute bottom-8 left-5 right-5">

              <div className="bg-black/60 backdrop-blur-md rounded-2xl p-4 text-center">

                {scannerError ? (
                  <p className="text-red-300 text-xs font-semibold">
                    {
                      scannerError
                    }
                  </p>
                ) : isScanning ? (
                  <p className="text-white text-xs font-semibold">
                    {
                      scannerStatus
                    }
                  </p>
                ) : (
                  <p className="text-white text-xs font-semibold">
                    Avvio fotocamera...
                  </p>
                )}

                <p className="text-white/60 text-[10px] mt-2">
                  Posiziona il codice
                  EAN-13 del libro
                  dentro il riquadro.
                </p>
              </div>
            </div>
          </div>

          <div className="bg-black px-5 pt-4 pb-[calc(env(safe-area-inset-bottom)+20px)]">
            <button
              type="button"
              onClick={
                closeScanner
              }
              className="w-full py-3.5 bg-white text-black rounded-2xl font-bold text-sm"
            >
              Chiudi fotocamera
            </button>
          </div>
        </div>
      )}

      {/* =================================================
          TAB BAR
      ================================================= */}

      <nav className="fixed bottom-0 left-0 right-0 z-30 bg-[#FBF9F5]/90 backdrop-blur-md border-t border-amber-900/10 flex justify-around py-2.5 max-w-lg mx-auto">

        <button
          onClick={() => {
            setActiveTab(
              'home'
            );
            setSelectedAuthor(
              null
            );
            setHomeSubView(
              'none'
            );
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab ===
            'home'
              ? 'text-amber-800'
              : 'text-amber-900/40'
          }`}
        >
          <Home className="w-5 h-5" />
          <span className="text-[10px] font-bold">
            Home
          </span>
        </button>

        <button
          onClick={() => {
            setActiveTab(
              'read'
            );
            setSelectedAuthor(
              null
            );
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab ===
            'read'
              ? 'text-amber-800'
              : 'text-amber-900/40'
          }`}
        >
          <BookOpen className="w-5 h-5" />
          <span className="text-[10px] font-bold">
            Letti
          </span>
        </button>

        <button
          onClick={() => {
            setActiveTab(
              'authors'
            );
            setSelectedAuthor(
              null
            );
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab ===
            'authors'
              ? 'text-amber-800'
              : 'text-amber-900/40'
          }`}
        >
          <Users className="w-5 h-5" />
          <span className="text-[10px] font-bold">
            Autori
          </span>
        </button>

        <button
          onClick={() => {
            setActiveTab(
              'settings'
            );
            setSelectedAuthor(
              null
            );
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab ===
            'settings'
              ? 'text-amber-800'
              : 'text-amber-900/40'
          }`}
        >
          <Settings className="w-5 h-5" />
          <span className="text-[10px] font-bold">
            Settings
          </span>
        </button>
      </nav>
    </div>
  );
}
