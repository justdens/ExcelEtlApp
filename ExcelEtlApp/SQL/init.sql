CREATE TABLE lokasi (
  id varchar PRIMARY KEY,
  lokasi varchar
);

CREATE TABLE trx (
  id serial PRIMARY KEY,
  bulan date,
  lokasi_id varchar,
  trx bigint
);

CREATE TABLE npp (
  id serial PRIMARY KEY,
  bulan date,
  lokasi_id varchar,
  npp bigint
);

CREATE OR REPLACE VIEW ticket_size AS
SELECT t.lokasi_id                 AS lokasiid,
       l.lokasi,
       t.bulan,
       t.trx,
       COALESCE(n.npp, 0::numeric) AS npp,
       CASE
           WHEN t.trx = 0::numeric THEN 0.0
           ELSE round(COALESCE(n.npp, 0::numeric) / t.trx, 4)
           END                     AS ticketsize
FROM (SELECT trx.lokasi_id,
             trx.bulan,
             sum(trx.trx) AS trx
      FROM trx
      GROUP BY trx.lokasi_id, trx.bulan) t
         LEFT JOIN (SELECT npp.lokasi_id,
                           npp.bulan,
                           sum(npp.npp) AS npp
                    FROM npp
                    GROUP BY npp.lokasi_id, npp.bulan) n ON t.lokasi_id = n.lokasi_id AND t.bulan = n.bulan
         JOIN lokasi l ON l.id = t.lokasi_id
