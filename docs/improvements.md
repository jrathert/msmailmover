# Possible performance improvements

The current script deliberately processes message moves one at a time. This is
simple, makes failures easy to identify, and creates only one active request at
a time. For large migrations, it also means one HTTP request per message move.

## Use a larger page size when listing messages

The script currently requests 100 messages per Graph page (`$top=100`). Microsoft
Graph can allow up to 1,000 messages per page for this endpoint. Increasing the
page size can reduce the number of *listing* requests substantially: 4,000
messages need roughly 40 pages at 100, or roughly 4 pages at 1,000.

This only reduces discovery requests. Each move would still be a separate
request. A conservative page size is useful when very large messages could make
responses slow or unreliable.

## Use Microsoft Graph JSON batching for moves

Microsoft Graph's `/$batch` endpoint accepts up to 20 independent operations
in one HTTP request. A future version could group up to 20 message `move`
operations into one batch. For example, 4,000 selected messages would need
about 200 HTTP requests instead of 4,000 move requests.

Batching reduces network round trips; it does **not** make Microsoft process
only 200 moves. Graph still accounts for every individual move, and individual
operations in a batch can be throttled or fail independently.

## Requirements for a safe batch implementation

A batch implementation should preserve the script's current safeguards:

1. Keep the existing collect-all-IDs phase before any messages are moved.
2. Submit batches sequentially rather than sending many batches concurrently.
3. Inspect every operation result inside each batch response.
4. Count only successful move responses as moved.
5. Retry only temporary failures, especially throttling responses, following
   their supplied retry delay.
6. Keep the final selected-versus-successfully-moved comparison and report any
   mismatch as an error.

This offers a large speed improvement without treating a batch response as an
all-or-nothing success.
