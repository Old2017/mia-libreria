'use client';

import React, {
  useCallback,
  useEffect,
  useRef,
  useState,
} from 'react';

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
  Camera,
  Loader2,
  AlertCircle,
} from 'lucide-react';

import * as XLSX from 'xlsx';

import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

import {
  BrowserMultiFormatReader,
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

  format:
    | 'cartaceo'
    | 'ebook'
    | 'ebook_and_paper';

  isRead: boolean;

  readMonth?: string;
  readYear?: number;
  readMonthYear?: string;

  rating?: number;

  notes?: string;

  isbn?: string;

  createdAt: number;
}


/* =========================================================
   COSTANTI
========================================================= */

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

const STORAGE_KEY = 'ios_library_books_v6';


/* =========================================================
   UTILITY ISBN
========================================================= */

function cleanISBN(value: string) {
  return value
    .replace(/[^0-9Xx]/g, '')
    .toUpperCase();
}

function isValidISBN(value: string) {
  const isbn = cleanISBN(value);

  return isbn.length === 10 || isbn.length === 13;
}


/* =========================================================
   COMPONENTE PRINCIPALE
========================================================= */

export default function LibraryApp() {
  const currentYearNum = new Date().getFullYear();

  /* =======================================================
     NAVIGAZIONE
  ======================================================= */

  const [activeTab, setActiveTab] = useState<
    'home' | 'read' | 'authors' | 'settings'
  >('home');

  /* =======================================================
     LIBRI
  ======================================================= */

  const [books, setBooks] = useState<BookItem[]>([]);
  const [isLoaded, setIsLoaded] = useState(false);

  /* =======================================================
     MODALI
  ======================================================= */

  const [isAddModalOpen, setIsAddModalOpen] =
    useState(false);

  const [selectedBookDetail, setSelectedBookDetail] =
    useState<BookItem | null>(null);

  const [selectedAuthor, setSelectedAuthor] =
    useState<string | null>(null);

  const [homeSubView, setHomeSubView] =
    useState<'none' | 'classics' | 'genres'>('none');

  const [selectedGenreHome, setSelectedGenreHome] =
    useState<string | null>(null);

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
     ISBN
  ======================================================= */

  const [isbnInput, setIsbnInput] =
    useState('');

  const [isSearchingIsbn, setIsSearchingIsbn] =
    useState(false);

  /* =======================================================
     SCANNER
  ======================================================= */

  const [isScannerOpen, setIsScannerOpen] =
    useState(false);

  const [scannerStatus, setScannerStatus] =
    useState('Inquadra il codice ISBN del libro');

  const [scannerError, setScannerError] =
    useState('');

  const scannerVideoRef =
    useRef<HTMLVideoElement | null>(null);

  const scannerControlsRef =
    useRef<{ stop: () => void } | null>(null);

  const scannerReaderRef =
    useRef<BrowserMultiFormatReader | null>(null);

  const scannerLockedRef =
    useRef(false);

  /* =======================================================
     DRAG
  ======================================================= */

  const [draggedIndex, setDraggedIndex] =
    useState<number | null>(null);

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
     ANNI
  ======================================================= */

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
        'Errore caricamento biblioteca:',
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
        'Errore salvataggio biblioteca:',
        error
      );
    }
  }, [books, isLoaded]);


  /* =======================================================
     RESET FORM
  ======================================================= */

  const resetForm = useCallback(() => {
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
  }, [currentYearNum]);


  /* =======================================================
     CERCA LIBRO VIA API NEXT.JS
  ======================================================= */

  const handleSearchBookByISBN = useCallback(
    async (code?: string) => {
      const raw = code ?? isbnInput;
      const isbn = cleanISBN(raw);

      if (!isbn) {
        alert('Inserisci un codice ISBN.');
        return;
      }

      if (!isValidISBN(isbn)) {
        alert(
          'Il codice inserito non sembra un ISBN valido.'
        );
        return;
      }

      setIsSearchingIsbn(true);

      try {
        const response = await fetch(
          `/api/books/isbn?isbn=${encodeURIComponent(
            isbn
          )}`,
          {
            method: 'GET',
            cache: 'no-store',
          }
        );

        let data: any = null;

        try {
          data = await response.json();
        } catch {
          throw new Error(
            'La risposta del server non è valida.'
          );
        }

        if (!response.ok || !data?.ok) {
          throw new Error(
            data?.error ||
              'Nessun libro trovato.'
          );
        }

        const book = data.book;

        setFormData((previous) => ({
          ...previous,

          isbn: book.isbn || isbn,

          title:
            book.title ||
            previous.title ||
            '',

          author:
            book.author ||
            previous.author ||
            '',

          publisher:
            book.publisher ||
            previous.publisher ||
            '',

          publishYear:
            book.publishYear ||
            previous.publishYear ||
            '',

          pages:
            book.pages ||
            previous.pages ||
            undefined,

          genre:
            book.genre ||
            previous.genre ||
            '',

          coverUrl:
            book.coverUrl ||
            previous.coverUrl ||
            '',
        }));

        setIsbnInput(
          book.isbn || isbn
        );

        setScannerStatus(
          'Libro trovato. Dati compilati automaticamente.'
        );

        return true;
      } catch (error) {
        console.error(
          'Errore ricerca ISBN:',
          error
        );

        alert(
          error instanceof Error
            ? error.message
            : 'Errore durante la ricerca del libro.'
        );

        return false;
      } finally {
        setIsSearchingIsbn(false);
      }
    },
    [isbnInput]
  );


  /* =======================================================
     STOP SCANNER
  ======================================================= */

  const stopScanner = useCallback(() => {
    try {
      scannerControlsRef.current?.stop();
    } catch (error) {
      console.warn(
        'Errore stop scanner:',
        error
      );
    }

    scannerControlsRef.current = null;

    scannerLockedRef.current = false;

    const video =
      scannerVideoRef.current;

    if (video?.srcObject) {
      const stream =
        video.srcObject as MediaStream;

      stream
        .getTracks()
        .forEach((track) => track.stop());

      video.srcObject = null;
    }

    scannerReaderRef.current = null;

    setIsScannerOpen(false);
  }, []);


  /* =======================================================
     AVVIO SCANNER
  ======================================================= */

  const startScanner = useCallback(
    async () => {
      setScannerError('');
      setScannerStatus(
        'Richiesta accesso alla fotocamera...'
      );

      scannerLockedRef.current = false;

      try {
        if (
          typeof window === 'undefined' ||
          !navigator.mediaDevices?.getUserMedia
        ) {
          throw new Error(
            'La fotocamera non è disponibile in questo browser.'
          );
        }

        if (!scannerVideoRef.current) {
          throw new Error(
            'Elemento video non disponibile.'
          );
        }

        /*
         * Hint per velocizzare la ricerca.
         * I libri normalmente utilizzano EAN-13
         * con prefisso 978/979.
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
            BarcodeFormat.ITF,
            BarcodeFormat.QR_CODE,
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

        scannerReaderRef.current = reader;

        const video =
          scannerVideoRef.current;

        /*
         * Importante:
         * NON usiamo reader.reset().
         *
         * La versione moderna di @zxing/browser
         * restituisce dei controls con stop().
         */
        const controls =
          await reader.decodeFromVideoDevice(
            undefined,
            video,
            async (
              result,
              error,
              controlsFromCallback
            ) => {
              if (!result) {
                return;
              }

              if (scannerLockedRef.current) {
                return;
              }

              scannerLockedRef.current = true;

              const rawText =
                result.getText();

              const detectedISBN =
                cleanISBN(rawText);

              console.log(
                'Codice rilevato:',
                rawText
              );

              setScannerStatus(
                `Codice rilevato: ${detectedISBN}`
              );

              /*
               * Fermiamo subito la scansione.
               */
              try {
                controlsFromCallback.stop();
              } catch {}

              try {
                controls.stop();
              } catch {}

              scannerControlsRef.current =
                null;

              /*
               * Chiudiamo il video.
               */
              if (video.srcObject) {
                const stream =
                  video.srcObject as MediaStream;

                stream
                  .getTracks()
                  .forEach((track) =>
                    track.stop()
                  );

                video.srcObject = null;
              }

              /*
               * Verifichiamo che sia compatibile
               * con un ISBN.
               */
              if (
                detectedISBN.length !== 13 &&
                detectedISBN.length !== 10
              ) {
                scannerLockedRef.current = false;

                setScannerStatus(
                  'Codice letto, ma non sembra un ISBN. Riprova.'
                );

                return;
              }

              setIsScannerOpen(false);

              setIsbnInput(
                detectedISBN
              );

              /*
               * Ora cerchiamo il libro.
               */
              await handleSearchBookByISBN(
                detectedISBN
              );

              scannerLockedRef.current =
                false;
            }
          );

        scannerControlsRef.current =
          controls;

        setScannerStatus(
          'Inquadra il codice a barre ISBN sul retro del libro'
        );
      } catch (error) {
        console.error(
          'Errore avvio scanner:',
          error
        );

        scannerControlsRef.current = null;

        if (
          error instanceof DOMException &&
          error.name === 'NotAllowedError'
        ) {
          setScannerError(
            'Accesso alla fotocamera negato. Vai in Impostazioni > Safari > Fotocamera e consenti l’accesso.'
          );
        } else if (
          error instanceof DOMException &&
          error.name === 'NotFoundError'
        ) {
          setScannerError(
            'Nessuna fotocamera disponibile.'
          );
        } else {
          setScannerError(
            error instanceof Error
              ? error.message
              : 'Impossibile avviare la fotocamera.'
          );
        }
      }
    },
    [handleSearchBookByISBN]
  );


  /* =======================================================
     APRI SCANNER
  ======================================================= */

  const openScanner = () => {
    setScannerError('');
    setScannerStatus(
      'Preparazione fotocamera...'
    );
    setIsScannerOpen(true);
  };


  /* =======================================================
     EFFECT SCANNER
  ======================================================= */

  useEffect(() => {
    if (!isScannerOpen) {
      return;
    }

    const timer = window.setTimeout(() => {
      startScanner();
    }, 150);

    return () => {
      window.clearTimeout(timer);
    };
  }, [isScannerOpen, startScanner]);


  /* =======================================================
     CLEANUP
  ======================================================= */

  useEffect(() => {
    return () => {
      try {
        scannerControlsRef.current?.stop();
      } catch {}

      const video =
        scannerVideoRef.current;

      if (video?.srcObject) {
        const stream =
          video.srcObject as MediaStream;

        stream
          .getTracks()
          .forEach((track) => track.stop());
      }
    };
  }, []);


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
        formData.publishCountry || '',

      coverUrl:
        formData.coverUrl || '',

      publisher:
        formData.publisher || '',

      publishYear:
        formData.publishYear || '',

      pages:
        Number(formData.pages) ||
        undefined,

      genre:
        formData.genre || '',

      seriesTag:
        formData.seriesTag || '',

      volume:
        formData.volume || '',

      isClassic:
        !!formData.isClassic,

      format:
        formData.format ||
        'cartaceo',

      isRead:
        !!formData.isRead,

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
        formData.notes || '',

      isbn:
        cleanISBN(
          formData.isbn || ''
        ),

      createdAt:
        formData.createdAt ||
        Date.now(),
    };

    if (formData.id) {
      setBooks((previous) =>
        previous.map((book) =>
          book.id === formData.id
            ? newBook
            : book
        )
      );
    } else {
      setBooks((previous) => [
        newBook,
        ...previous,
      ]);
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

    setBooks((previous) =>
      previous.filter(
        (book) => book.id !== id
      )
    );

    if (
      selectedBookDetail?.id === id
    ) {
      setSelectedBookDetail(null);
    }
  };


  /* =======================================================
     STATISTICHE
  ======================================================= */

  const totalBooks =
    books.length;

  const totalCartacei =
    books.filter(
      (book) =>
        book.format === 'cartaceo' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const totalEbook =
    books.filter(
      (book) =>
        book.format === 'ebook' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readBooks =
    books.filter(
      (book) => book.isRead
    );

  const readBooksCount =
    readBooks.length;

  const readCartacei =
    readBooks.filter(
      (book) =>
        book.format === 'cartaceo' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readEbook =
    readBooks.filter(
      (book) =>
        book.format === 'ebook' ||
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

      return !!book.readMonthYear?.includes(
        currentYearNum.toString()
      );
    });

  const readThisYearCount =
    readThisYearBooks.length;

  const readThisYearCartacei =
    readThisYearBooks.filter(
      (book) =>
        book.format === 'cartaceo' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readThisYearEbook =
    readThisYearBooks.filter(
      (book) =>
        book.format === 'ebook' ||
        book.format ===
          'ebook_and_paper'
    ).length;


  /* =======================================================
     FILTRI
  ======================================================= */

  const availableGenres =
    Array.from(
      new Set(
        books
          .map((book) => book.genre)
          .filter(Boolean)
      )
    );

  const availableYears =
    Array.from(
      new Set(
        books
          .map((book) => {
            if (book.readYear) {
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
    );

  const readBooksFiltered =
    books
      .filter(
        (book) => book.isRead
      )
      .filter((book) =>
        filterGenre === 'all'
          ? true
          : book.genre ===
            filterGenre
      )
      .filter((book) => {
        if (
          filterFormat === 'all'
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
          filterYear === 'all'
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
        if (!searchQuery) {
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
        const author =
          book.author.trim() ||
          'Autore Sconosciuto';

        if (
          !accumulator[author]
        ) {
          accumulator[author] =
            [];
        }

        accumulator[author].push(
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
        const genre =
          book.genre?.trim() ||
          'Generico / Altro';

        if (
          !accumulator[genre]
        ) {
          accumulator[genre] =
            [];
        }

        accumulator[genre].push(
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


  /* =======================================================
     DRAG
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
      readOnlyBooks[
        draggedIndex
      ];

    const targetBook =
      readOnlyBooks[index];

    if (
      !itemToMove ||
      !targetBook
    ) {
      return;
    }

    const updated = [
      ...books,
    ];

    const sourceIndex =
      updated.findIndex(
        (book) =>
          book.id ===
          itemToMove.id
      );

    const targetIndex =
      updated.findIndex(
        (book) =>
          book.id ===
          targetBook.id
      );

    if (
      sourceIndex < 0 ||
      targetIndex < 0
    ) {
      return;
    }

    updated.splice(
      sourceIndex,
      1
    );

    updated.splice(
      targetIndex,
      0,
      itemToMove
    );

    setDraggedIndex(index);
    setBooks(updated);
  };

  const handleDragEnd = () => {
    setDraggedIndex(null);
  };


  /* =======================================================
     FORMAT
  ======================================================= */

  const formatLabel = (
    format: string
  ) => {
    if (
      format === 'cartaceo'
    ) {
      return 'Cartaceo';
    }

    if (
      format === 'ebook'
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
     AUTORE / LIBRI
  ======================================================= */

  const renderAuthorGroup = (
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
            !accumulator[author]
          ) {
            accumulator[author] =
              [];
          }

          accumulator[author].push(
            book
          );

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
      .map((author) => (
        <div
          key={author}
          className="space-y-2 pt-2"
        >
          <h3 className="font-serif font-bold text-sm text-amber-950 border-b border-amber-900/10 pb-1">
            {author}
          </h3>

          <div className="grid grid-cols-2 gap-3">
            {map[author]
              .sort((a, b) =>
                (
                  a.volume || ''
                ).localeCompare(
                  b.volume || ''
                )
              )
              .map((book) => (
                <div
                  key={book.id}
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
              ))}
          </div>
        </div>
      ));
  };


  /* =======================================================
     EXPORT EXCEL
  ======================================================= */

  const exportToExcel = () => {
    const data =
      books.map(
        (book) => ({
          Titolo:
            book.title,
          Autore:
            book.author,
          'Paese di Pubblicazione':
            book.publishCountry ||
            '-',
          Classico:
            book.isClassic
              ? 'Sì'
              : 'No',
          Stato:
            book.isRead
              ? 'Letto'
              : 'In Biblioteca',
          Formato:
            formatLabel(
              book.format
            ),
          Editore:
            book.publisher ||
            '',
          'Anno Pubblicazione':
            book.publishYear ||
            '',
          Genere:
            book.genre ||
            '',
          'Serie / Tag':
            book.seriesTag ||
            '',
          Volume:
            book.volume ||
            '',
          Pagine:
            book.pages ||
            '',
          'Mese e Anno di Lettura':
            book.readMonthYear ||
            '',
          ISBN:
            book.isbn ||
            '',
        })
      );

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
      books.map(
        (book) => [
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
          formatLabel(
            book.format
          ),
          book.genre ||
            '-',
          book.readMonthYear ||
            '-',
        ]
      );

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
    const dataStr =
      'data:text/json;charset=utf-8,' +
      encodeURIComponent(
        JSON.stringify(
          books,
          null,
          2
        )
      );

    const anchor =
      document.createElement(
        'a'
      );

    anchor.setAttribute(
      'href',
      dataStr
    );

    anchor.setAttribute(
      'download',
      'backup_libreria.json'
    );

    document.body.appendChild(
      anchor
    );

    anchor.click();

    anchor.remove();
  };


  /* =======================================================
     IMPORT BACKUP
  ======================================================= */

  const handleImportBackup = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file =
      event.target.files?.[0];

    if (!file) return;

    const reader =
      new FileReader();

    reader.onload = (
      loadEvent
    ) => {
      try {
        const parsed =
          JSON.parse(
            loadEvent.target
              ?.result as string
          );

        if (
          !Array.isArray(
            parsed
          )
        ) {
          throw new Error(
            'Formato non valido'
          );
        }

        setBooks(parsed);

        alert(
          'Backup ripristinato con successo!'
        );
      } catch (error) {
        console.error(error);

        alert(
          'File di backup non valido.'
        );
      }
    };

    reader.readAsText(
      file,
      'UTF-8'
    );

    event.target.value = '';
  };


  /* =======================================================
     CLEAR
  ======================================================= */

  const handleClearAll = () => {
    if (
      !confirm(
        'ATTENZIONE: verranno cancellati TUTTI i libri. Procedere?'
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
            {activeTab === 'home' &&
              'La Mia Collezione'}

            {activeTab === 'read' &&
              'Cronologia Letture'}

            {activeTab === 'authors' &&
              'Catalogo Autori'}

            {activeTab === 'settings' &&
              'Gestione Dati'}
          </span>

          <h1 className="text-2xl font-serif font-extrabold tracking-tight text-amber-950">
            {activeTab === 'home' &&
              'Home'}

            {activeTab === 'read' &&
              'Libri Letti'}

            {activeTab === 'authors' &&
              (selectedAuthor ||
                'Autori')}

            {activeTab ===
              'settings' &&
              'Impostazioni'}
          </h1>
        </div>

        {activeTab === 'read' && (
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

      {activeTab === 'home' && (
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
                        Generi nella
                        Biblioteca
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
              <div className="bg-[#FFFDF9] rounded-3xl p-5 shadow-sm border border-amber-900/10">

                <div className="flex items-center gap-4">

                  <div className="w-12 h-12 rounded-2xl bg-amber-800 text-amber-50 flex items-center justify-center">
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

                <div className="pt-3 mt-3 border-t border-amber-900/10 text-center text-xs font-semibold">
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

                <div className="bg-[#FFFDF9] p-4 rounded-3xl shadow-sm border border-amber-900/10">

                  <BookOpen className="w-5 h-5 mb-6" />

                  <span className="text-3xl font-serif font-black">
                    {
                      totalBooks
                    }
                  </span>

                  <span className="text-xs font-bold text-amber-800/60 uppercase tracking-wider block mt-1">
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


                <div className="bg-[#FFFDF9] p-4 rounded-3xl shadow-sm border border-amber-900/10">

                  <CheckCircle2 className="w-5 h-5 mb-6 text-emerald-700" />

                  <span className="text-3xl font-serif font-black">
                    {
                      readBooksCount
                    }
                  </span>

                  <span className="text-xs font-bold text-amber-800/60 uppercase tracking-wider block mt-1">
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


              <div className="space-y-3 pt-2">

                <h2 className="text-xs font-bold text-amber-800/60 uppercase tracking-wider">
                  Esplora Categorie
                </h2>

                <div className="grid grid-cols-2 gap-3.5">

                  <div
                    onClick={() =>
                      setHomeSubView(
                        'classics'
                      )
                    }
                    className="bg-gradient-to-br from-amber-700 to-amber-900 text-amber-50 p-4 rounded-3xl shadow-md cursor-pointer h-32 flex flex-col justify-between"
                  >
                    <Sparkles />

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
                        libri conservati
                      </span>
                    </div>
                  </div>


                  <div
                    onClick={() =>
                      setHomeSubView(
                        'genres'
                      )
                    }
                    className="bg-gradient-to-br from-stone-800 to-amber-950 text-amber-50 p-4 rounded-3xl shadow-md cursor-pointer h-32 flex flex-col justify-between"
                  >
                    <Layers />

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
                  </div>

                </div>
              </div>
            </>
          )}
        </div>
      )}


      {/* =================================================
          LETTI
      ================================================= */}

      {activeTab === 'read' && (
        <div className="p-4 space-y-4 max-w-lg mx-auto">

          <div className="relative">
            <Search className="w-4 h-4 absolute left-3.5 top-3 text-amber-800/40" />

            <input
              type="text"
              placeholder="Cerca nei libri letti..."
              value={
                searchQuery
              }
              onChange={(event) =>
                setSearchQuery(
                  event.target.value
                )
              }
              className="w-full pl-9 pr-4 py-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-2xl text-xs"
            />
          </div>


          <div className="flex gap-2 overflow-x-auto pb-1">

            <select
              value={
                filterGenre
              }
              onChange={(event) =>
                setFilterGenre(
                  event.target.value
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
                    key={
                      genre
                    }
                    value={
                      genre
                    }
                  >
                    {
                      genre
                    }
                  </option>
                )
              )}
            </select>


            <select
              value={
                filterFormat
              }
              onChange={(event) =>
                setFilterFormat(
                  event.target.value
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
              onChange={(event) =>
                setFilterYear(
                  event.target.value
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
                    key={
                      year
                    }
                    value={
                      year || ''
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


          <p className="text-[11px] text-amber-800/50">
            Trascina i libri per riordinarli.
          </p>


          <div className="space-y-3">

            {readBooksFiltered.length ===
            0 ? (
              <div className="text-center py-12 bg-[#FFFDF9] rounded-3xl border border-dashed border-amber-900/20">
                <BookOpen className="w-8 h-8 mx-auto mb-2 text-amber-800/30" />

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
                      event
                    ) =>
                      handleDragOver(
                        event,
                        index
                      )
                    }
                    onDragEnd={
                      handleDragEnd
                    }
                    className={`bg-[#FFFDF9] p-3 rounded-2xl border border-amber-900/10 flex items-center gap-3 ${
                      draggedIndex ===
                      index
                        ? 'bg-amber-100'
                        : ''
                    }`}
                  >

                    <GripVertical className="w-5 h-5 text-amber-800/30" />

                    <div
                      onClick={() =>
                        setSelectedBookDetail(
                          book
                        )
                      }
                      className="w-13 h-19 bg-amber-100/40 rounded-lg overflow-hidden flex-shrink-0 cursor-pointer"
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
                          <Book className="w-5 h-5" />
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

                      <div className="flex text-amber-500 mt-1">
                        {[
                          0,
                          1,
                          2,
                          3,
                          4,
                        ].map(
                          (
                            star
                          ) => (
                            <Star
                              key={
                                star
                              }
                              className={`w-3 h-3 ${
                                star <
                                (book.rating ||
                                  0)
                                  ? 'fill-amber-500'
                                  : 'text-amber-200'
                              }`}
                            />
                          )
                        )}
                      </div>

                      <p className="text-[10px] mt-1">
                        Letto:{' '}
                        <b>
                          {
                            book.readMonthYear ||
                            'Data non specificata'
                          }
                        </b>
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

      {activeTab === 'authors' && (
        <div className="p-4 max-w-lg mx-auto">

          {!selectedAuthor ? (
            <div className="space-y-2">

              {sortedAuthors.length ===
              0 ? (
                <div className="text-center py-12">
                  <Users className="w-8 h-8 mx-auto mb-2" />

                  <p className="text-xs">
                    Nessun autore
                    presente.
                  </p>
                </div>
              ) : (
                sortedAuthors.map(
                  (author) => (
                    <div
                      key={
                        author
                      }
                      onClick={() =>
                        setSelectedAuthor(
                          author
                        )
                      }
                      className="bg-[#FFFDF9] p-4 rounded-2xl shadow-sm border border-amber-900/10 flex justify-between items-center cursor-pointer"
                    >
                      <div>
                        <h3 className="font-serif font-bold">
                          {
                            author
                          }
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

                      <ChevronRight />
                    </div>
                  )
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
                ← Torna alla lista
              </button>

              <h2 className="text-lg font-serif font-bold">
                {
                  selectedAuthor
                }
              </h2>

              <div className="grid grid-cols-2 gap-3">

                {authorsMap[
                  selectedAuthor
                ]?.map(
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
                      className="bg-[#FFFDF9] p-3 rounded-2xl border border-amber-900/10 cursor-pointer"
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
                            <Book />
                          </div>
                        )}
                      </div>

                      <h4 className="font-serif font-bold text-xs">
                        {
                          book.title
                        }
                      </h4>

                      <p className="text-[10px] text-amber-800/60">
                        {
                          book.publishYear
                        }
                      </p>
                    </div>
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

      {activeTab === 'settings' && (
        <div className="p-4 space-y-5 max-w-lg mx-auto">

          <div className="bg-[#FFFDF9] rounded-3xl p-4 border border-amber-900/10 space-y-3">

            <h2 className="text-xs font-bold uppercase tracking-wider">
              Esporta
            </h2>

            <button
              onClick={
                exportToExcel
              }
              className="w-full p-3 bg-emerald-50 text-emerald-900 rounded-2xl font-bold text-xs flex items-center gap-3"
            >
              <FileSpreadsheet />
              Esporta Excel
            </button>

            <button
              onClick={
                exportToPDF
              }
              className="w-full p-3 bg-rose-50 text-rose-900 rounded-2xl font-bold text-xs flex items-center gap-3"
            >
              <FileText />
              Esporta PDF
            </button>
          </div>


          <div className="bg-[#FFFDF9] rounded-3xl p-4 border border-amber-900/10 space-y-3">

            <h2 className="text-xs font-bold uppercase tracking-wider">
              Backup
            </h2>

            <button
              onClick={
                exportBackup
              }
              className="w-full p-3 bg-amber-100 rounded-2xl font-bold text-xs flex items-center gap-3"
            >
              <Download />
              Salva Backup JSON
            </button>

            <label className="w-full p-3 bg-stone-100 rounded-2xl font-bold text-xs flex items-center gap-3 cursor-pointer">
              <Upload />
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


          <div className="bg-[#FFFDF9] rounded-3xl p-4 border border-rose-200">

            <button
              onClick={
                handleClearAll
              }
              className="w-full p-3 bg-rose-700 text-white rounded-2xl font-bold text-xs flex items-center justify-center gap-2"
            >
              <Trash2 className="w-4 h-4" />
              Cancella Intera
              Biblioteca
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

            <div className="flex justify-between">

              <h2 className="text-lg font-serif font-bold">
                Scheda Libro
              </h2>

              <button
                onClick={() =>
                  setSelectedBookDetail(
                    null
                  )
                }
              >
                <X />
              </button>
            </div>


            <div className="flex gap-4">

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

                <h2 className="font-serif font-bold">
                  {
                    selectedBookDetail.title
                  }
                </h2>

                <p className="text-xs text-amber-800/70 mt-1">
                  {
                    selectedBookDetail.author
                  }
                </p>

                {selectedBookDetail.isbn && (
                  <p className="text-[10px] mt-2">
                    ISBN:{' '}
                    {
                      selectedBookDetail.isbn
                    }
                  </p>
                )}

                <div className="flex flex-wrap gap-2 mt-3">

                  <span className="text-[10px] px-2 py-1 rounded-full bg-amber-100 font-bold">
                    {formatLabel(
                      selectedBookDetail.format
                    )}
                  </span>

                  {selectedBookDetail.isRead && (
                    <span className="text-[10px] px-2 py-1 rounded-full bg-emerald-100 text-emerald-900 font-bold">
                      Letto
                    </span>
                  )}
                </div>
              </div>
            </div>


            <div className="grid grid-cols-2 gap-3 bg-amber-50 p-4 rounded-2xl text-xs">

              <div>
                <span className="text-[10px] text-amber-800/50 block">
                  Editore
                </span>

                <b>
                  {
                    selectedBookDetail.publisher ||
                    '-'
                  }
                </b>
              </div>

              <div>
                <span className="text-[10px] text-amber-800/50 block">
                  Anno
                </span>

                <b>
                  {
                    selectedBookDetail.publishYear ||
                    '-'
                  }
                </b>
              </div>

              <div>
                <span className="text-[10px] text-amber-800/50 block">
                  Pagine
                </span>

                <b>
                  {
                    selectedBookDetail.pages ||
                    '-'
                  }
                </b>
              </div>

              <div>
                <span className="text-[10px] text-amber-800/50 block">
                  Genere
                </span>

                <b>
                  {
                    selectedBookDetail.genre ||
                    '-'
                  }
                </b>
              </div>

              <div>
                <span className="text-[10px] text-amber-800/50 block">
                  Serie
                </span>

                <b>
                  {
                    selectedBookDetail.seriesTag ||
                    '-'
                  }
                </b>
              </div>

              <div>
                <span className="text-[10px] text-amber-800/50 block">
                  Volume
                </span>

                <b>
                  {
                    selectedBookDetail.volume ||
                    '-'
                  }
                </b>
              </div>

              {selectedBookDetail.isRead && (
                <div className="col-span-2 border-t border-amber-900/10 pt-2">
                  <span className="text-[10px] text-amber-800/50 block">
                    Data lettura
                  </span>

                  <b>
                    {
                      selectedBookDetail.readMonthYear ||
                      '-'
                    }
                  </b>
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
          MODALE AGGIUNTA
      ================================================= */}

      {isAddModalOpen && (
        <div className="fixed inset-0 z-50 bg-amber-950/40 backdrop-blur-sm flex items-end sm:items-center justify-center">

          <div className="bg-[#FFFDF9] w-full max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[94vh] overflow-y-auto p-6 space-y-4">

            <div className="flex justify-between items-center border-b border-amber-900/10 pb-3">

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
                <X />
              </button>
            </div>


            {/* ===========================================
                SCANNER
            =========================================== */}

            <div className="bg-amber-100/50 p-4 rounded-2xl border border-amber-900/10 space-y-3">

              <div className="flex justify-between items-center">

                <div>
                  <span className="text-[10px] font-bold uppercase tracking-wider">
                    Scansione ISBN
                  </span>

                  <p className="text-[10px] text-amber-800/60 mt-1">
                    Usa la fotocamera
                    posteriore del
                    dispositivo.
                  </p>
                </div>

                <Scan className="w-5 h-5" />
              </div>


              <button
                type="button"
                onClick={
                  openScanner
                }
                className="w-full py-3 bg-amber-800 text-white rounded-xl font-bold text-xs flex items-center justify-center gap-2 active:scale-95"
              >
                <Camera className="w-4 h-4" />

                Scansiona con
                Fotocamera
              </button>


              <div className="text-center text-[10px] text-amber-800/50">
                oppure inserisci
                l'ISBN manualmente
              </div>


              <div className="flex gap-2">

                <input
                  type="text"
                  inputMode="numeric"
                  autoComplete="off"
                  placeholder="978..."
                  value={
                    isbnInput
                  }
                  onChange={(
                    event
                  ) => {
                    const value =
                      event.target
                        .value;

                    setIsbnInput(
                      value
                    );

                    setFormData(
                      (
                        previous
                      ) => ({
                        ...previous,
                        isbn: cleanISBN(
                          value
                        ),
                      })
                    );
                  }}
                  className="flex-1 p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl text-xs"
                />

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
            </div>


            {/* ===========================================
                FORM
            =========================================== */}

            <form
              onSubmit={
                handleSaveBook
              }
              className="space-y-3 text-xs"
            >

              <input
                type="text"
                placeholder="Titolo *"
                value={
                  formData.title ||
                  ''
                }
                onChange={(
                  event
                ) =>
                  setFormData(
                    {
                      ...formData,
                      title:
                        event.target
                          .value,
                    }
                  )
                }
                className="w-full p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl"
                required
              />


              <div className="grid grid-cols-2 gap-2">

                <input
                  type="text"
                  placeholder="Autore *"
                  value={
                    formData.author ||
                    ''
                  }
                  onChange={(
                    event
                  ) =>
                    setFormData(
                      {
                        ...formData,
                        author:
                          event.target
                            .value,
                      }
                    )
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl"
                  required
                />

                <input
                  type="text"
                  placeholder="Paese"
                  value={
                    formData.publishCountry ||
                    ''
                  }
                  onChange={(
                    event
                  ) =>
                    setFormData(
                      {
                        ...formData,
                        publishCountry:
                          event
                            .target
                            .value,
                      }
                    )
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl"
                />
              </div>


              <input
                type="text"
                placeholder="URL Copertina"
                value={
                  formData.coverUrl ||
                  ''
                }
                onChange={(
                  event
                ) =>
                  setFormData(
                    {
                      ...formData,
                      coverUrl:
                        event.target
                          .value,
                    }
                  )
                }
                className="w-full p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl"
              />


              <div className="grid grid-cols-2 gap-2">

                <input
                  type="text"
                  placeholder="Editore"
                  value={
                    formData.publisher ||
                    ''
                  }
                  onChange={(
                    event
                  ) =>
                    setFormData(
                      {
                        ...formData,
                        publisher:
                          event
                            .target
                            .value,
                      }
                    )
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl"
                />

                <input
                  type="text"
                  placeholder="Anno"
                  value={
                    formData.publishYear ||
                    ''
                  }
                  onChange={(
                    event
                  ) =>
                    setFormData(
                      {
                        ...formData,
                        publishYear:
                          event
                            .target
                            .value,
                      }
                    )
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl"
                />
              </div>


              <div className="grid grid-cols-2 gap-2">

                <input
                  type="text"
                  placeholder="Genere"
                  value={
                    formData.genre ||
                    ''
                  }
                  onChange={(
                    event
                  ) =>
                    setFormData(
                      {
                        ...formData,
                        genre:
                          event
                            .target
                            .value,
                      }
                    )
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl"
                />

                <input
                  type="text"
                  inputMode="numeric"
                  placeholder="Pagine"
                  value={
                    formData.pages ||
                    ''
                  }
                  onChange={(
                    event
                  ) => {
                    const value =
                      event.target.value.replace(
                        /\D/g,
                        ''
                      );

                    setFormData({
                      ...formData,
                      pages: value
                        ? parseInt(
                            value,
                            10
                          )
                        : undefined,
                    });
                  }}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl"
                />
              </div>


              <div className="grid grid-cols-2 gap-2">

                <input
                  type="text"
                  placeholder="Serie / Tag"
                  value={
                    formData.seriesTag ||
                    ''
                  }
                  onChange={(
                    event
                  ) =>
                    setFormData(
                      {
                        ...formData,
                        seriesTag:
                          event
                            .target
                            .value,
                      }
                    )
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl"
                />

                <input
                  type="text"
                  placeholder="Volume"
                  value={
                    formData.volume ||
                    ''
                  }
                  onChange={(
                    event
                  ) =>
                    setFormData(
                      {
                        ...formData,
                        volume:
                          event
                            .target
                            .value,
                      }
                    )
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl"
                />
              </div>


              {/* CLASSICO */}

              <label className="flex items-center justify-between p-3 bg-amber-50 rounded-xl border border-amber-900/10 cursor-pointer">

                <span className="font-bold flex items-center gap-2">
                  <Sparkles className="w-4 h-4 text-amber-600" />
                  È un Classico?
                </span>

                <input
                  type="checkbox"
                  checked={
                    !!formData.isClassic
                  }
                  onChange={(
                    event
                  ) =>
                    setFormData(
                      {
                        ...formData,
                        isClassic:
                          event
                            .target
                            .checked,
                      }
                    )
                  }
                  className="w-5 h-5 accent-amber-800"
                />
              </label>


              {/* FORMATO */}

              <div className="p-3 bg-amber-50 rounded-xl border border-amber-900/10 space-y-2">

                <span className="font-bold">
                  Formato
                </span>

                <div className="flex gap-2">

                  <button
                    type="button"
                    onClick={() =>
                      setFormData(
                        {
                          ...formData,
                          format:
                            'cartaceo',
                        }
                      )
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
                      setFormData(
                        {
                          ...formData,
                          format:
                            'ebook',
                        }
                      )
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
                  <label className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      checked={
                        formData.format ===
                        'ebook_and_paper'
                      }
                      onChange={(
                        event
                      ) =>
                        setFormData(
                          {
                            ...formData,
                            format:
                              event
                                .target
                                .checked
                                ? 'ebook_and_paper'
                                : 'ebook',
                          }
                        )
                      }
                    />

                    Acquistato anche
                    in formato
                    Cartaceo
                  </label>
                )}
              </div>


              {/* LETTO */}

              <label className="flex items-center justify-between p-3 bg-amber-50 rounded-xl border border-amber-900/10 cursor-pointer">

                <span className="font-bold">
                  Letto
                </span>

                <input
                  type="checkbox"
                  checked={
                    !!formData.isRead
                  }
                  onChange={(
                    event
                  ) =>
                    setFormData(
                      {
                        ...formData,
                        isRead:
                          event
                            .target
                            .checked,
                      }
                    )
                  }
                  className="w-5 h-5 accent-amber-800"
                />
              </label>


              {/* DATA LETTURA */}

              {formData.isRead && (
                <div className="p-3 bg-amber-100/40 rounded-xl space-y-2">

                  <span className="font-bold">
                    Data di Lettura
                  </span>

                  <div className="grid grid-cols-2 gap-2">

                    <select
                      value={
                        formData.readMonth ||
                        'Gennaio'
                      }
                      onChange={(
                        event
                      ) =>
                        setFormData(
                          {
                            ...formData,
                            readMonth:
                              event
                                .target
                                .value,
                          }
                        )
                      }
                      className="p-3 bg-white border rounded-xl"
                    >
                      {MONTHS.map(
                        (
                          month
                        ) => (
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
                      onChange={(
                        event
                      ) =>
                        setFormData(
                          {
                            ...formData,
                            readYear:
                              parseInt(
                                event
                                  .target
                                  .value,
                                10
                              ),
                          }
                        )
                      }
                      className="p-3 bg-white border rounded-xl"
                    >
                      {yearsList.map(
                        (
                          year
                        ) => (
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
                className="w-full py-3 bg-amber-800 text-white rounded-xl font-bold shadow-md active:scale-95"
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
          MODALE FOTOCAMERA
      ================================================= */}

      {isScannerOpen && (
        <div className="fixed inset-0 z-[100] bg-black flex flex-col">

          <div className="flex items-center justify-between p-4 text-white">

            <div>
              <h2 className="font-serif font-bold text-lg">
                Scansiona ISBN
              </h2>

              <p className="text-xs text-white/60">
                Posiziona il codice
                al centro
              </p>
            </div>

            <button
              onClick={
                stopScanner
              }
              className="w-10 h-10 rounded-full bg-white/10 flex items-center justify-center"
            >
              <X />
            </button>
          </div>


          <div className="flex-1 flex items-center justify-center relative overflow-hidden">

            <video
              ref={
                scannerVideoRef
              }
              autoPlay
              muted
              playsInline
              className="w-full h-full object-cover"
            />


            {/* RETTANGOLO DI SCANSIONE */}

            <div className="absolute inset-x-8 top-1/2 -translate-y-1/2 h-32 border-2 border-white rounded-2xl shadow-[0_0_0_9999px_rgba(0,0,0,0.45)]">

              <div className="absolute left-0 right-0 top-1/2 h-0.5 bg-red-500 shadow-lg" />
            </div>


            {/* MESSAGGIO */}

            <div className="absolute bottom-8 left-4 right-4">

              <div className="bg-black/70 backdrop-blur-md rounded-2xl p-4 text-white text-center">

                {scannerError ? (
                  <div className="text-red-300 text-xs flex items-center justify-center gap-2">
                    <AlertCircle className="w-4 h-4" />

                    {scannerError}
                  </div>
                ) : (
                  <div className="text-xs flex items-center justify-center gap-2">

                    {!scannerLockedRef.current && (
                      <Loader2 className="w-4 h-4 animate-spin" />
                    )}

                    {
                      scannerStatus
                    }
                  </div>
                )}

              </div>
            </div>
          </div>


          <div className="p-5 bg-black text-white text-center">

            <p className="text-[11px] text-white/60">
              Suggerimento: usa la
              fotocamera posteriore e
              avvicinati lentamente al
              codice a barre.
            </p>

            {scannerError && (
              <button
                onClick={() => {
                  setScannerError(
                    ''
                  );
                  startScanner();
                }}
                className="mt-3 px-5 py-2.5 bg-white text-black rounded-xl text-xs font-bold"
              >
                Riprova
              </button>
            )}
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
            activeTab === 'home'
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
            activeTab === 'read'
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
            activeTab === 'authors'
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
            activeTab === 'settings'
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
