import { NextResponse } from 'next/server';

function cleanISBN(value: string) {
  return value
    .replace(/[^0-9Xx]/g, '')
    .toUpperCase();
}

export async function GET(request: Request) {
  try {
    const url = new URL(request.url);
    const rawISBN = url.searchParams.get('isbn') || '';

    const isbn = cleanISBN(rawISBN);

    if (!isbn) {
      return NextResponse.json(
        {
          ok: false,
          error: 'ISBN mancante.',
        },
        { status: 400 }
      );
    }

    if (isbn.length !== 10 && isbn.length !== 13) {
      return NextResponse.json(
        {
          ok: false,
          error: 'ISBN non valido.',
        },
        { status: 400 }
      );
    }

    console.log('Ricerca ISBN:', isbn);

    const openLibraryUrl =
      `https://openlibrary.org/isbn/${encodeURIComponent(isbn)}.json`;

    const response = await fetch(openLibraryUrl, {
      method: 'GET',
      headers: {
        Accept: 'application/json',
      },
      cache: 'no-store',
    });

    if (response.status === 404) {
      return NextResponse.json(
        {
          ok: false,
          error: `Nessun libro trovato per ISBN ${isbn}.`,
        },
        { status: 404 }
      );
    }

    if (!response.ok) {
      console.error(
        'Open Library HTTP error:',
        response.status
      );

      return NextResponse.json(
        {
          ok: false,
          error:
            'Il servizio di ricerca libri non è al momento disponibile.',
        },
        { status: 502 }
      );
    }

    const data = await response.json();

    console.log('Risposta Open Library:', data);

    const title =
      data.title ||
      '';

    const authors =
      Array.isArray(data.authors)
        ? data.authors
        : [];

    let author = '';

    if (authors.length > 0) {
      author = authors
        .map((item: any) => {
          if (
            typeof item === 'string'
          ) {
            return item;
          }

          return (
            item?.name ||
            ''
          );
        })
        .filter(Boolean)
        .join(', ');
    }

    const publishers =
      Array.isArray(data.publishers)
        ? data.publishers
        : [];

    const publisher =
      publishers.length > 0
        ? typeof publishers[0] === 'string'
          ? publishers[0]
          : publishers[0]?.name || ''
        : '';

    let publishYear = '';

    if (data.publish_date) {
      const match =
        String(data.publish_date).match(
          /\d{4}/
        );

      if (match) {
        publishYear = match[0];
      }
    }

    const pages =
      typeof data.number_of_pages ===
      'number'
        ? data.number_of_pages
        : undefined;

    let coverUrl = '';

    if (data.covers?.[0]) {
      coverUrl =
        `https://covers.openlibrary.org/b/id/${data.covers[0]}-L.jpg`;
    }

    let genre = '';

    if (
      Array.isArray(
        data.subjects
      ) &&
      data.subjects.length > 0
    ) {
      const firstSubject =
        data.subjects[0];

      if (
        typeof firstSubject ===
        'string'
      ) {
        genre = firstSubject;
      } else {
        genre =
          firstSubject?.name ||
          '';
      }
    }

    const book = {
      isbn,
      title,
      author,
      publisher,
      publishYear,
      pages,
      genre,
      coverUrl,
    };

    if (!title) {
      return NextResponse.json(
        {
          ok: false,
          error:
            'Il libro è stato trovato, ma non contiene un titolo utilizzabile.',
        },
        { status: 404 }
      );
    }

    return NextResponse.json({
      ok: true,
      book,
    });
  } catch (error) {
    console.error(
      'Errore API ISBN:',
      error
    );

    return NextResponse.json(
      {
        ok: false,
        error:
          'Errore durante la ricerca del libro.',
      },
      { status: 500 }
    );
  }
}
