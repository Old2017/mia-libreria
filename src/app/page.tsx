'use client';

import React, { useState, useEffect, useRef } from 'react';
import {
  Home,
  BookOpen,
  Users,
  Settings,
  Plus,
  Search,
  Scan,
  Filter,
  Trash2,
  Download,
  Upload,
  CheckCircle2,
  Book,
  Smartphone,
  ChevronRight,
  Star,
  X,
  FileSpreadsheet,
  FileText,
  RotateCcw
} from 'lucide-react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

// --- INTERFACCIA DATI LIBRO ---
export interface BookItem {
  id: string;
  title: string;
  author: string;
  coverUrl?: string;
  publisher?: string;
  publishYear?: string;
  pages?: number;
  genre?: string;
  seriesTag?: string;
  volumeOrder?: number;
  format: 'cartaceo' | 'ebook';
  isRead: boolean;
  readMonthYear?: string; // Es: "2026-05" o "Maggio 2026"
  readYearOnly?: number; // Es: 2026
  rating?: number;
  notes?: string;
  isbn?: string;
  createdAt: number;
}

export default function LibraryApp() {
  const [activeTab, setActiveTab] = useState<'home' | 'read' | 'authors' | 'settings'>('home');
  const [books, setBooks] = useState<BookItem[]>([]);
  const [isLoaded, setIsLoaded] = useState(false);

  // Modal e Form
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [selectedBookDetail, setSelectedBookDetail] = useState<BookItem | null>(null);
  const [selectedAuthor, setSelectedAuthor] = useState<string | null>(null);

  // Scanner ISBN
  const [isScannerOpen, setIsScannerOpen] = useState(false);
  const [isbnInput, setIsbnInput] = useState('');
  const [isSearchingIsbn, setIsSearchingIsbn] = useState(false);

  // Filtri Sezione Letti
  const [filterGenre, setFilterGenre] = useState<string>('all');
  const [filterFormat, setFilterFormat] = useState<string>('all');
  const [filterYear, setFilterYear] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState('');

  // Form State
  const [formData, setFormData] = useState<Partial<BookItem>>({
    title: '',
    author: '',
    coverUrl: '',
    publisher: '',
    publishYear: '',
    pages: 0,
    genre: '',
    seriesTag: '',
    volumeOrder: 1,
    format: 'cartaceo',
    isRead: false,
    readMonthYear: '',
    rating: 5,
    notes: '',
  });

  // Caricamento Dati
  useEffect(() => {
    const saved = localStorage.getItem('ios_library_books_v2');
    if (saved) {
      try {
        setBooks(JSON.parse(saved));
      } catch (e) {
        console.error('Errore caricamento storage', e);
      }
    }
    setIsLoaded(true);
  }, []);

  // Salvataggio Dati
  useEffect(() => {
    if (isLoaded) {
      localStorage.setItem('ios_library_books_v2', JSON.stringify(books));
    }
  }, [books, isLoaded]);

  // Statistiche Home
  const totalBooks = books.length;
  const readBooksCount = books.filter((b) => b.isRead).length;
  const currentYear = new Date().getFullYear();
  const readThisYearCount = books.filter((b) => {
    if (!b.isRead) return false;
    if (b.readYearOnly) return b.readYearOnly === currentYear;
    if (b.readMonthYear && b.readMonthYear.includes(currentYear.toString())) return true;
    return false;
  }).length;

  // Ricerca Google Books per ISBN / Titolo
  const handleSearchBookByISBN = async (codeToSearch?: string) => {
    const query = codeToSearch || isbnInput;
    if (!query) return;
    setIsSearchingIsbn(true);

    try {
      const res = await fetch(`https://www.googleapis.com/books/v1/volumes?q=isbn:${query}`);
      const data = await res.json();

      if (data.items && data.items.length > 0) {
        const info = data.items[0].volumeInfo;
        setFormData((prev) => ({
          ...prev,
          title: info.title || '',
          author: info.authors ? info.authors.join(', ') : '',
          publisher: info.publisher || '',
          publishYear: info.publishedDate ? info.publishedDate.substring(0, 4) : '',
          pages: info.pageCount || 0,
          genre: info.categories ? info.categories[0] : '',
          coverUrl: info.imageLinks?.thumbnail?.replace('http:', 'https:') || '',
          isbn: query,
        }));
        setIsScannerOpen(false);
      } else {
        alert('Nessun libro trovato con questo codice. Inserisci i dati manualmente.');
      }
    } catch (err) {
      console.error(err);
      alert('Errore durante il recupero dei dati del libro.');
    } finally {
      setIsSearchingIsbn(false);
    }
  };

  // Salva nuovo libro o modifica
  const handleSaveBook = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.title || !formData.author) {
      alert('Inserisci almeno Titolo e Autore.');
      return;
    }

    let yearNum: number | undefined = undefined;
    if (formData.readMonthYear) {
      const match = formData.readMonthYear.match(/\d{4}/);
      if (match) yearNum = parseInt(match[0]);
    }

    const newBook: BookItem = {
      id: formData.id || Date.now().toString(),
      title: formData.title || '',
      author: formData.author || '',
      coverUrl: formData.coverUrl || '',
      publisher: formData.publisher || '',
      publishYear: formData.publishYear || '',
      pages: Number(formData.pages) || 0,
      genre: formData.genre || '',
      seriesTag: formData.seriesTag || '',
      volumeOrder: Number(formData.volumeOrder) || 1,
      format: formData.format || 'cartaceo',
      isRead: !!formData.isRead,
      readMonthYear: formData.readMonthYear || '',
      readYearOnly: yearNum,
      rating: Number(formData.rating) || 5,
      notes: formData.notes || '',
      isbn: formData.isbn || '',
      createdAt: formData.createdAt || Date.now(),
    };

    if (formData.id) {
      setBooks(books.map((b) => (b.id === formData.id ? newBook : b)));
    } else {
      setBooks([newBook, ...books]);
    }

    setIsAddModalOpen(false);
    resetForm();
  };

  const resetForm = () => {
    setFormData({
      title: '',
      author: '',
      coverUrl: '',
      publisher: '',
      publishYear: '',
      pages: 0,
      genre: '',
      seriesTag: '',
      volumeOrder: 1,
      format: 'cartaceo',
      isRead: false,
      readMonthYear: '',
      rating: 5,
      notes: '',
    });
  };

  const handleDeleteBook = (id: string) => {
    if (confirm('Sei sicuro di voler eliminare questo libro?')) {
      setBooks(books.filter((b) => b.id !== id));
      if (selectedBookDetail?.id === id) setSelectedBookDetail(null);
    }
  };

  // Esportazione Excel
  const exportToExcel = () => {
    const dataToExport = books.map((b) => ({
      Titolo: b.title,
      Autore: b.author,
      Stato: b.isRead ? 'Letto' : 'Da Leggere / Posseduto',
      Formato: b.format,
      Editore: b.publisher,
      'Anno Pubblicazione': b.publishYear,
      Genere: b.genre,
      'Serie / Tag': b.seriesTag,
      'Volume N°': b.volumeOrder,
      Pagine: b.pages,
      'Mese e Anno di Lettura': b.readMonthYear,
      Valutazione: b.rating ? `${b.rating}/5` : '',
      ISBN: b.isbn,
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Biblioteca');
    XLSX.writeFile(workbook, 'La_Mia_Biblioteca.xlsx');
  };

  // Esportazione PDF
  const exportToPDF = () => {
    const doc = new jsPDF();
    doc.text('La Mia Biblioteca - Report', 14, 15);

    const tableData = books.map((b) => [
      b.title,
      b.author,
      b.isRead ? 'Letto' : 'In Libreria',
      b.format,
      b.genre || '-',
      b.readMonthYear || '-',
    ]);

    autoTable(doc, {
      head: [['Titolo', 'Autore', 'Stato', 'Formato', 'Genere', 'Data Lettura']],
      body: tableData,
      startY: 20,
    });

    doc.save('La_Mia_Biblioteca.pdf');
  };

  // Esportazione JSON Backup
  const exportBackup = () => {
    const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(books));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute('href', dataStr);
    downloadAnchor.setAttribute('download', 'backup_libreria.json');
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
  };

  // Importazione JSON Backup
  const handleImportBackup = (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileReader = new FileReader();
    if (e.target.files && e.target.files[0]) {
      fileReader.readAsText(e.target.files[0], 'UTF-8');
      fileReader.onload = (event) => {
        try {
          const parsed = JSON.parse(event.target?.result as string);
          if (Array.isArray(parsed)) {
            setBooks(parsed);
            alert('Backup ripristinato con successo!');
          }
        } catch (err) {
          alert('File di backup non valido.');
        }
      };
    }
  };

  // Reset Totale
  const handleClearAll = () => {
    if (confirm('ATTENZIONE: Verranno cancellati TUTTI i libri salvati. Procedere?')) {
      setBooks([]);
      localStorage.removeItem('ios_library_books_v2');
    }
  };

  // Libri Letti Filtri
  const readBooksFiltered = books
    .filter((b) => b.isRead)
    .filter((b) => (filterGenre === 'all' ? true : b.genre === filterGenre))
    .filter((b) => (filterFormat === 'all' ? true : b.format === filterFormat))
    .filter((b) => {
      if (filterYear === 'all') return true;
      return b.readMonthYear?.includes(filterYear) || b.readYearOnly?.toString() === filterYear;
    })
    .filter((b) =>
      searchQuery === ''
        ? true
        : b.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
          b.author.toLowerCase().includes(searchQuery.toLowerCase())
    );

  // Raggruppamento Autori
  const authorsMap = books.reduce((acc, book) => {
    const authorName = book.author.trim() || 'Autore Sconosciuto';
    if (!acc[authorName]) acc[authorName] = [];
    acc[authorName].push(book);
    return acc;
  }, {} as Record<string, BookItem[]>);

  const sortedAuthors = Object.keys(authorsMap).sort((a, b) => a.localeCompare(b));

  // Generi e Anni per Filtri
  const availableGenres = Array.from(new Set(books.map((b) => b.genre).filter(Boolean)));
  const availableYears = Array.from(
    new Set(
      books
        .map((b) => {
          if (b.readYearOnly) return b.readYearOnly.toString();
          const match = b.readMonthYear?.match(/\d{4}/);
          return match ? match[0] : null;
        })
        .filter(Boolean)
    )
  );

  return (
    <div className="min-h-screen bg-[#F2F2F7] text-slate-900 font-sans pb-24 select-none">
      {/* Header Principale */}
      <header className="sticky top-0 z-20 bg-[#F2F2F7]/80 backdrop-blur-md border-b border-slate-200/60 px-4 py-3 flex justify-between items-center">
        <h1 className="text-2xl font-extrabold tracking-tight text-black">
          {activeTab === 'home' && 'Home'}
          {activeTab === 'read' && 'Libri Letti'}
          {activeTab === 'authors' && (selectedAuthor ? selectedAuthor : 'Autori')}
          {activeTab === 'settings' && 'Impostazioni'}
        </h1>
        <button
          onClick={() => {
            resetForm();
            setIsAddModalOpen(true);
          }}
          className="w-9 h-9 bg-blue-600 text-white rounded-full flex items-center justify-center shadow-md active:scale-95 transition-transform"
        >
          <Plus className="w-5 h-5" />
        </button>
      </header>

      {/* VISTA 1: HOME */}
      {activeTab === 'home' && (
        <div className="p-4 space-y-6 max-w-lg mx-auto">
          {/* Panoramica Statistiche */}
          <div className="grid grid-cols-2 gap-4">
            <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 flex flex-col items-start">
              <div className="w-10 h-10 rounded-xl bg-blue-50 flex items-center justify-center text-blue-600 mb-3">
                <BookOpen className="w-5 h-5" />
              </div>
              <span className="text-3xl font-bold text-slate-900">{totalBooks}</span>
              <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider mt-1">
                Nella Biblioteca
              </span>
            </div>

            <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 flex flex-col items-start">
              <div className="w-10 h-10 rounded-xl bg-emerald-50 flex items-center justify-center text-emerald-600 mb-3">
                <CheckCircle2 className="w-5 h-5" />
              </div>
              <span className="text-3xl font-bold text-slate-900">{readBooksCount}</span>
              <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider mt-1">
                Libri Letti Totali
              </span>
            </div>
          </div>

          {/* Panoramica Anno in Corso */}
          <div className="bg-gradient-to-br from-blue-600 to-indigo-700 rounded-3xl p-6 text-white shadow-lg relative overflow-hidden">
            <div className="relative z-10">
              <span className="text-xs uppercase font-bold tracking-widest text-blue-200">
                Anno In Corso ({currentYear})
              </span>
              <div className="flex items-baseline gap-2 mt-2">
                <span className="text-5xl font-black">{readThisYearCount}</span>
                <span className="text-lg text-blue-100 font-medium">libri letti</span>
              </div>
              <p className="text-xs text-blue-200/80 mt-3">
                Continua così! La tua collezione continua a crescere.
              </p>
            </div>
          </div>

          {/* Ultimi Libri Aggiunti */}
          <div className="space-y-3">
            <h2 className="text-lg font-bold text-slate-800">Aggiunti di recente</h2>
            <div className="space-y-2">
              {books.slice(0, 4).map((book) => (
                <div
                  key={book.id}
                  onClick={() => setSelectedBookDetail(book)}
                  className="bg-white p-3 rounded-2xl shadow-sm border border-slate-100 flex items-center gap-3 active:bg-slate-50 transition-colors cursor-pointer"
                >
                  <div className="w-12 h-16 bg-slate-100 rounded-lg overflow-hidden flex-shrink-0 relative border border-slate-200">
                    {book.coverUrl ? (
                      <img src={book.coverUrl} alt={book.title} className="w-full h-full object-cover" />
                    ) : (
                      <div className="w-full h-full flex items-center justify-center text-slate-400">
                        <Book className="w-6 h-6" />
                      </div>
                    )}
                  </div>
                  <div className="flex-1 min-w-0">
                    <h3 className="font-bold text-sm text-slate-900 truncate">{book.title}</h3>
                    <p className="text-xs text-slate-500 truncate">{book.author}</p>
                    <div className="flex items-center gap-2 mt-1">
                      <span className="text-[10px] px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 font-medium">
                        {book.format === 'cartaceo' ? '📖 Cartaceo' : '📱 eBook'}
                      </span>
                      {book.isRead && (
                        <span className="text-[10px] px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-600 font-semibold">
                          Letto
                        </span>
                      )}
                    </div>
                  </div>
                  <ChevronRight className="w-5 h-5 text-slate-300" />
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* VISTA 2: LIBRI LETTI (Cronologico) */}
      {activeTab === 'read' && (
        <div className="p-4 space-y-4 max-w-lg mx-auto">
          {/* Barra di Ricerca e Filtri */}
          <div className="space-y-3">
            <div className="relative">
              <Search className="w-4 h-4 absolute left-3 top-3 text-slate-400" />
              <input
                type="text"
                placeholder="Cerca nei libri letti..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-9 pr-4 py-2.5 bg-white border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>

            {/* Select Filtri */}
            <div className="flex gap-2 overflow-x-auto pb-1 text-xs">
              <select
                value={filterGenre}
                onChange={(e) => setFilterGenre(e.target.value)}
                className="bg-white border border-slate-200 rounded-xl px-3 py-2 font-medium text-slate-700"
              >
                <option value="all">Tutti i generi</option>
                {availableGenres.map((g) => (
                  <option key={g} value={g}>
                    {g}
                  </option>
                ))}
              </select>

              <select
                value={filterFormat}
                onChange={(e) => setFilterFormat(e.target.value)}
                className="bg-white border border-slate-200 rounded-xl px-3 py-2 font-medium text-slate-700"
              >
                <option value="all">Tutti i formati</option>
                <option value="cartaceo">Cartaceo</option>
                <option value="ebook">eBook</option>
              </select>

              <select
                value={filterYear}
                onChange={(e) => setFilterYear(e.target.value)}
                className="bg-white border border-slate-200 rounded-xl px-3 py-2 font-medium text-slate-700"
              >
                <option value="all">Tutti gli anni</option>
                {availableYears.map((y) => (
                  <option key={y} value={y!}>
                    {y}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Elenco Libri Letti */}
          <div className="space-y-3">
            {readBooksFiltered.length === 0 ? (
              <div className="text-center py-12 bg-white rounded-2xl border border-dashed border-slate-200">
                <BookOpen className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                <p className="text-sm text-slate-500 font-medium">Nessun libro letto trovato.</p>
              </div>
            ) : (
              readBooksFiltered.map((book) => (
                <div
                  key={book.id}
                  onClick={() => setSelectedBookDetail(book)}
                  className="bg-white p-3 rounded-2xl shadow-sm border border-slate-100 flex items-center gap-3 active:bg-slate-50 transition-colors cursor-pointer"
                >
                  <div className="w-14 h-20 bg-slate-100 rounded-lg overflow-hidden flex-shrink-0 relative border border-slate-200">
                    {book.coverUrl ? (
                      <img src={book.coverUrl} alt={book.title} className="w-full h-full object-cover" />
                    ) : (
                      <div className="w-full h-full flex items-center justify-center text-slate-400">
                        <Book className="w-6 h-6" />
                      </div>
                    )}
                    <div className="absolute top-1 right-1 bg-emerald-500 text-white rounded-full p-0.5 shadow">
                      <CheckCircle2 className="w-3 h-3" />
                    </div>
                  </div>

                  <div className="flex-1 min-w-0">
                    <h3 className="font-bold text-sm text-slate-900 truncate">{book.title}</h3>
                    <p className="text-xs text-slate-500 truncate">{book.author}</p>
                    <div className="flex items-center gap-1.5 mt-1">
                      {book.rating && (
                        <div className="flex text-amber-400">
                          {[...Array(5)].map((_, i) => (
                            <Star
                              key={i}
                              className={`w-3 h-3 ${
                                i < book.rating! ? 'fill-amber-400' : 'text-slate-200'
                              }`}
                            />
                          ))}
                        </div>
                      )}
                    </div>
                    <p className="text-[11px] text-slate-400 mt-1">
                      Letto: <span className="font-medium text-slate-600">{book.readMonthYear || 'Data non specificata'}</span>
                    </p>
                  </div>
                  <ChevronRight className="w-5 h-5 text-slate-300" />
                </div>
              ))
            )}
          </div>
        </div>
      )}

      {/* VISTA 3: AUTORI (Alfabetico con copertine) */}
      {activeTab === 'authors' && (
        <div className="p-4 space-y-4 max-w-lg mx-auto">
          {!selectedAuthor ? (
            /* Lista Autori Alfabetica */
            <div className="space-y-2">
              {sortedAuthors.length === 0 ? (
                <div className="text-center py-12 bg-white rounded-2xl border border-dashed border-slate-200">
                  <Users className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                  <p className="text-sm text-slate-500 font-medium">Nessun autore presente.</p>
                </div>
              ) : (
                sortedAuthors.map((author) => {
                  const authorBooks = authorsMap[author];
                  return (
                    <div
                      key={author}
                      onClick={() => setSelectedAuthor(author)}
                      className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 flex justify-between items-center active:bg-slate-50 transition-colors cursor-pointer"
                    >
                      <div>
                        <h3 className="font-bold text-base text-slate-900">{author}</h3>
                        <p className="text-xs text-slate-500 font-medium">
                          {authorBooks.length} {authorBooks.length === 1 ? 'libro posseduto' : 'libri posseduti'}
                        </p>
                      </div>
                      <ChevronRight className="w-5 h-5 text-slate-300" />
                    </div>
                  );
                })
              )}
            </div>
          ) : (
            /* Dettaglio Autore Selezionato (Libri ordinati per volume/uscita) */
            <div className="space-y-4">
              <button
                onClick={() => setSelectedAuthor(null)}
                className="text-xs font-semibold text-blue-600 flex items-center gap-1 mb-2"
              >
                ← Torna alla lista autori
              </button>

              <div className="grid grid-cols-2 gap-4">
                {authorsMap[selectedAuthor]
                  ?.sort((a, b) => (a.volumeOrder || 0) - (b.volumeOrder || 0))
                  .map((book) => (
                    <div
                      key={book.id}
                      onClick={() => setSelectedBookDetail(book)}
                      className="bg-white p-3 rounded-2xl shadow-sm border border-slate-100 flex flex-col cursor-pointer relative group"
                    >
                      {/* Copertina con Segno di Spunta e Formato */}
                      <div className="w-full h-44 bg-slate-100 rounded-xl overflow-hidden relative border border-slate-200 mb-2">
                        {book.coverUrl ? (
                          <img src={book.coverUrl} alt={book.title} className="w-full h-full object-cover" />
                        ) : (
                          <div className="w-full h-full flex items-center justify-center text-slate-400">
                            <Book className="w-8 h-8" />
                          </div>
                        )}

                        {/* Spunta Verde se Letto */}
                        {book.isRead && (
                          <div className="absolute top-2 right-2 bg-emerald-500 text-white p-1 rounded-full shadow-md">
                            <CheckCircle2 className="w-4 h-4" />
                          </div>
                        )}

                        {/* Badge Formato */}
                        <div className="absolute bottom-2 left-2 bg-black/60 backdrop-blur-md text-white text-[10px] font-semibold px-2 py-0.5 rounded-full">
                          {book.format === 'cartaceo' ? '📖 Cartaceo' : '📱 eBook'}
                        </div>

                        {/* Badge Numero Volume */}
                        {book.volumeOrder && (
                          <div className="absolute top-2 left-2 bg-blue-600 text-white text-[10px] font-bold px-2 py-0.5 rounded-full shadow">
                            Vol. {book.volumeOrder}
                          </div>
                        )}
                      </div>

                      <h4 className="font-bold text-xs text-slate-900 line-clamp-2">{book.title}</h4>
                      <p className="text-[11px] text-slate-500 mt-1">
                        {book.publishYear ? `Anno: ${book.publishYear}` : ''}
                      </p>
                    </div>
                  ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* VISTA 4: SETTINGS */}
      {activeTab === 'settings' && (
        <div className="p-4 space-y-6 max-w-lg mx-auto">
          {/* Sezione Esportazione Documenti */}
          <div className="bg-white rounded-2xl p-4 shadow-sm border border-slate-100 space-y-3">
            <h2 className="text-xs font-bold text-slate-400 uppercase tracking-wider">
              Esporta la tua Biblioteca
            </h2>

            <button
              onClick={exportToExcel}
              className="w-full p-3 bg-emerald-50 text-emerald-700 rounded-xl font-semibold text-sm flex items-center gap-3 hover:bg-emerald-100 transition-colors"
            >
              <FileSpreadsheet className="w-5 h-5" /> Esporta in Foglio Excel (.xlsx)
            </button>

            <button
              onClick={exportToPDF}
              className="w-full p-3 bg-red-50 text-red-700 rounded-xl font-semibold text-sm flex items-center gap-3 hover:bg-red-100 transition-colors"
            >
              <FileText className="w-5 h-5" /> Esporta Report PDF
            </button>
          </div>

          {/* Backup e Ripristino */}
          <div className="bg-white rounded-2xl p-4 shadow-sm border border-slate-100 space-y-3">
            <h2 className="text-xs font-bold text-slate-400 uppercase tracking-wider">
              Backup & Ripristino Dati
            </h2>

            <button
              onClick={exportBackup}
              className="w-full p-3 bg-blue-50 text-blue-700 rounded-xl font-semibold text-sm flex items-center gap-3 hover:bg-blue-100 transition-colors"
            >
              <Download className="w-5 h-5" /> Salva Backup Dati (JSON)
            </button>

            <label className="w-full p-3 bg-slate-50 text-slate-700 rounded-xl font-semibold text-sm flex items-center gap-3 hover:bg-slate-100 transition-colors cursor-pointer">
              <Upload className="w-5 h-5" /> Ripristina Backup da File
              <input type="file" accept=".json" onChange={handleImportBackup} className="hidden" />
            </label>
          </div>

          {/* Reset e Gestione Dati */}
          <div className="bg-white rounded-2xl p-4 shadow-sm border border-slate-100 space-y-3">
            <h2 className="text-xs font-bold text-red-400 uppercase tracking-wider">Zona Pericolo</h2>
            <button
              onClick={handleClearAll}
              className="w-full p-3 bg-red-500 text-white rounded-xl font-semibold text-sm flex items-center justify-center gap-2 hover:bg-red-600 transition-colors"
            >
              <Trash2 className="w-4 h-4" /> Cancella Intera Biblioteca
            </button>
          </div>
        </div>
      )}

      {/* DETTAGLIO SCHEDA LIBRO (MODALE iOS STYLE) */}
      {selectedBookDetail && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">
          <div className="bg-white w-full max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[90vh] overflow-y-auto p-6 space-y-4 shadow-2xl relative">
            <button
              onClick={() => setSelectedBookDetail(null)}
              className="absolute top-4 right-4 p-2 bg-slate-100 rounded-full text-slate-500"
            >
              <X className="w-5 h-5" />
            </button>

            <div className="flex gap-4 items-start">
              <div className="w-24 h-36 bg-slate-100 rounded-xl overflow-hidden border border-slate-200 flex-shrink-0">
                {selectedBookDetail.coverUrl ? (
                  <img
                    src={selectedBookDetail.coverUrl}
                    alt={selectedBookDetail.title}
                    className="w-full h-full object-cover"
                  />
                ) : (
                  <div className="w-full h-full flex items-center justify-center text-slate-400">
                    <Book className="w-8 h-8" />
                  </div>
                )}
              </div>

              <div className="flex-1">
                <h2 className="text-xl font-bold text-slate-900">{selectedBookDetail.title}</h2>
                <p className="text-sm font-medium text-slate-500">{selectedBookDetail.author}</p>
                <div className="flex gap-2 mt-2">
                  <span className="text-xs px-2.5 py-1 rounded-full bg-blue-50 text-blue-600 font-semibold">
                    {selectedBookDetail.format === 'cartaceo' ? '📖 Cartaceo' : '📱 eBook'}
                  </span>
                  {selectedBookDetail.isRead ? (
                    <span className="text-xs px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-600 font-semibold">
                      Letto
                    </span>
                  ) : (
                    <span className="text-xs px-2.5 py-1 rounded-full bg-amber-50 text-amber-600 font-semibold">
                      In Biblioteca
                    </span>
                  )}
                </div>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3 bg-slate-50 p-4 rounded-2xl text-xs text-slate-600">
              <div>
                <span className="block text-slate-400 font-medium">Editore:</span>
                <span className="font-semibold text-slate-800">{selectedBookDetail.publisher || '-'}</span>
              </div>
              <div>
                <span className="block text-slate-400 font-medium">Anno Pubblicazione:</span>
                <span className="font-semibold text-slate-800">{selectedBookDetail.publishYear || '-'}</span>
              </div>
              <div>
                <span className="block text-slate-400 font-medium">Pagine:</span>
                <span className="font-semibold text-slate-800">{selectedBookDetail.pages || '-'}</span>
              </div>
              <div>
                <span className="block text-slate-400 font-medium">Genere:</span>
                <span className="font-semibold text-slate-800">{selectedBookDetail.genre || '-'}</span>
              </div>
              <div>
                <span className="block text-slate-400 font-medium">Serie / Tag:</span>
                <span className="font-semibold text-slate-800">{selectedBookDetail.seriesTag || '-'}</span>
              </div>
              <div>
                <span className="block text-slate-400 font-medium">Volume N°:</span>
                <span className="font-semibold text-slate-800">{selectedBookDetail.volumeOrder || '-'}</span>
              </div>
              {selectedBookDetail.isRead && (
                <div className="col-span-2 border-t border-slate-200 pt-2">
                  <span className="block text-slate-400 font-medium">Mese & Anno di Lettura:</span>
                  <span className="font-semibold text-slate-800">{selectedBookDetail.readMonthYear || '-'}</span>
                </div>
              )}
            </div>

            <button
              onClick={() => handleDeleteBook(selectedBookDetail.id)}
              className="w-full py-3 bg-red-50 text-red-600 rounded-xl font-bold text-sm flex items-center justify-center gap-2"
            >
              <Trash2 className="w-4 h-4" /> Elimina Libro
            </button>
          </div>
        </div>
      )}

      {/* MODALE AGGIUNTA / SCANSIONE LIBRO */}
      {isAddModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">
          <div className="bg-white w-full max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[90vh] overflow-y-auto p-6 space-y-4 shadow-2xl relative">
            <div className="flex justify-between items-center border-b border-slate-100 pb-3">
              <h2 className="text-lg font-bold text-slate-900">Aggiungi Libro</h2>
              <button onClick={() => setIsAddModalOpen(false)} className="p-2 text-slate-400">
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Pulsante / Sezione Scanner Codice a Barre */}
            <div className="bg-blue-50 p-4 rounded-2xl space-y-3">
              <div className="flex justify-between items-center">
                <span className="text-xs font-bold text-blue-900 uppercase">Cerca con Barcode / ISBN</span>
                <Scan className="w-5 h-5 text-blue-600" />
              </div>
              <div className="flex gap-2">
                <input
                  type="text"
                  placeholder="Inserisci o scansiona codice ISBN"
                  value={isbnInput}
                  onChange={(e) => setIsbnInput(e.target.value)}
                  className="flex-1 p-2.5 bg-white border border-blue-200 rounded-xl text-sm focus:outline-none"
                />
                <button
                  type="button"
                  onClick={() => handleSearchBookByISBN()}
                  disabled={isSearchingIsbn}
                  className="px-4 bg-blue-600 text-white rounded-xl text-xs font-bold"
                >
                  {isSearchingIsbn ? '...' : 'Cerca'}
                </button>
              </div>
            </div>

            {/* Form Dati Libro */}
            <form onSubmit={handleSaveBook} className="space-y-3 text-sm">
              <input
                type="text"
                placeholder="Titolo *"
                value={formData.title}
                onChange={(e) => setFormData({ ...formData, title: e.target.value })}
                className="w-full p-3 border border-slate-200 rounded-xl"
                required
              />
              <input
                type="text"
                placeholder="Autore *"
                value={formData.author}
                onChange={(e) => setFormData({ ...formData, author: e.target.value })}
                className="w-full p-3 border border-slate-200 rounded-xl"
                required
              />
              <input
                type="text"
                placeholder="URL Copertina Immagine"
                value={formData.coverUrl}
                onChange={(e) => setFormData({ ...formData, coverUrl: e.target.value })}
                className="w-full p-3 border border-slate-200 rounded-xl"
              />

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Editore"
                  value={formData.publisher}
                  onChange={(e) => setFormData({ ...formData, publisher: e.target.value })}
                  className="p-3 border border-slate-200 rounded-xl"
                />
                <input
                  type="text"
                  placeholder="Anno Pubblicazione"
                  value={formData.publishYear}
                  onChange={(e) => setFormData({ ...formData, publishYear: e.target.value })}
                  className="p-3 border border-slate-200 rounded-xl"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Genere"
                  value={formData.genre}
                  onChange={(e) => setFormData({ ...formData, genre: e.target.value })}
                  className="p-3 border border-slate-200 rounded-xl"
                />
                <input
                  type="number"
                  placeholder="Pagine"
                  value={formData.pages || ''}
                  onChange={(e) => setFormData({ ...formData, pages: parseInt(e.target.value) })}
                  className="p-3 border border-slate-200 rounded-xl"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Serie / Tag"
                  value={formData.seriesTag}
                  onChange={(e) => setFormData({ ...formData, seriesTag: e.target.value })}
                  className="p-3 border border-slate-200 rounded-xl"
                />
                <input
                  type="number"
                  placeholder="Volume N° (uscita)"
                  value={formData.volumeOrder || 1}
                  onChange={(e) => setFormData({ ...formData, volumeOrder: parseInt(e.target.value) })}
                  className="p-3 border border-slate-200 rounded-xl"
                />
              </div>

              <div className="flex items-center justify-between p-3 bg-slate-50 rounded-xl">
                <span className="font-semibold text-slate-700">Formato:</span>
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() => setFormData({ ...formData, format: 'cartaceo' })}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold ${
                      formData.format === 'cartaceo' ? 'bg-blue-600 text-white' : 'bg-slate-200 text-slate-600'
                    }`}
                  >
                    Cartaceo
                  </button>
                  <button
                    type="button"
                    onClick={() => setFormData({ ...formData, format: 'ebook' })}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold ${
                      formData.format === 'ebook' ? 'bg-blue-600 text-white' : 'bg-slate-200 text-slate-600'
                    }`}
                  >
                    eBook
                  </button>
                </div>
              </div>

              <div className="flex items-center justify-between p-3 bg-slate-50 rounded-xl">
                <span className="font-semibold text-slate-700">L'hai già letto?</span>
                <input
                  type="checkbox"
                  checked={formData.isRead}
                  onChange={(e) => setFormData({ ...formData, isRead: e.target.checked })}
                  className="w-5 h-5 accent-blue-600"
                />
              </div>

              {formData.isRead && (
                <input
                  type="text"
                  placeholder="Mese e Anno di lettura (es: Maggio 2026)"
                  value={formData.readMonthYear}
                  onChange={(e) => setFormData({ ...formData, readMonthYear: e.target.value })}
                  className="w-full p-3 border border-slate-200 rounded-xl"
                />
              )}

              <button
                type="submit"
                className="w-full py-3 bg-blue-600 text-white rounded-xl font-bold text-sm shadow-md"
              >
                Salva Libro
              </button>
            </form>
          </div>
        </div>
      )}

      {/* TAB BAR INFERIORE (Stile iOS Nativo) */}
      <nav className="fixed bottom-0 left-0 right-0 z-30 bg-white/80 backdrop-blur-lg border-t border-slate-200 flex justify-around py-2 max-w-lg mx-auto">
        <button
          onClick={() => {
            setActiveTab('home');
            setSelectedAuthor(null);
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab === 'home' ? 'text-blue-600' : 'text-slate-400'
          }`}
        >
          <Home className="w-5 h-5" />
          <span className="text-[10px] font-medium">Home</span>
        </button>

        <button
          onClick={() => {
            setActiveTab('read');
            setSelectedAuthor(null);
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab === 'read' ? 'text-blue-600' : 'text-slate-400'
          }`}
        >
          <BookOpen className="w-5 h-5" />
          <span className="text-[10px] font-medium">Letti</span>
        </button>

        <button
          onClick={() => {
            setActiveTab('authors');
            setSelectedAuthor(null);
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab === 'authors' ? 'text-blue-600' : 'text-slate-400'
          }`}
        >
          <Users className="w-5 h-5" />
          <span className="text-[10px] font-medium">Autori</span>
        </button>

        <button
          onClick={() => {
            setActiveTab('settings');
            setSelectedAuthor(null);
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab === 'settings' ? 'text-blue-600' : 'text-slate-400'
          }`}
        >
          <Settings className="w-5 h-5" />
          <span className="text-[10px] font-medium">Settings</span>
        </button>
      </nav>
    </div>
  );
}