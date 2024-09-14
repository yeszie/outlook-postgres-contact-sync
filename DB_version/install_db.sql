CREATE TABLE contacts (
    id SERIAL PRIMARY KEY,
    first_name VARCHAR(255),
    last_name VARCHAR(255),
    email VARCHAR(255) UNIQUE NOT NULL,
    company VARCHAR(255),
    business_phone VARCHAR(50),
    mobile_phone VARCHAR(50),
    last_contact TIMESTAMP,
    created_at TIMESTAMP DEFAULT NOW()
);

CREATE TABLE calendar_events (
    id SERIAL PRIMARY KEY,
    subject VARCHAR(255),
    start_time TIMESTAMP,
    end_time TIMESTAMP,
    organizer VARCHAR(255),
    attendees TEXT,
    last_interaction TIMESTAMP,
    created_at TIMESTAMP DEFAULT NOW()
);

CREATE TABLE contact_history (
    id SERIAL PRIMARY KEY,
    contact_id INT REFERENCES contacts(id),
    change_time TIMESTAMP DEFAULT NOW(),
    changed_by VARCHAR(255),
    old_data JSONB,
    new_data JSONB
);

CREATE TABLE calendar_event_history (
    id SERIAL PRIMARY KEY,
    event_id INT REFERENCES calendar_events(id),
    change_time TIMESTAMP DEFAULT NOW(),
    changed_by VARCHAR(255),
    old_data JSONB,
    new_data JSONB
);

CREATE TABLE blacklist (
    id SERIAL PRIMARY KEY,
    email VARCHAR(255) UNIQUE NOT NULL
);

-- Przykład dodania wpisu do blacklisty:
-- INSERT INTO blacklist (email) VALUES ('spam@example.com');

ALTER TABLE contacts
    ALTER COLUMN business_phone TYPE VARCHAR(255),
    ALTER COLUMN mobile_phone TYPE VARCHAR(255);

ALTER TABLE blacklist ADD COLUMN domain VARCHAR(255);
ALTER TABLE blacklist ALTER COLUMN email DROP NOT NULL;

ALTER TABLE blacklist ADD COLUMN prefix VARCHAR(255);
INSERT INTO blacklist (prefix) VALUES ('noreply@*');
INSERT INTO blacklist (prefix) VALUES ('no-reply@*');





ALTER TABLE contact_history
DROP CONSTRAINT contact_history_contact_id_fkey;

ALTER TABLE contact_history
ADD CONSTRAINT contact_history_contact_id_fkey
FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE CASCADE;
ALTER TABLE contacts ADD COLUMN notes TEXT;
ALTER TABLE contacts ADD COLUMN title VARCHAR(100);
ALTER TABLE contacts ADD COLUMN middle_name VARCHAR(100);
ALTER TABLE contacts ADD COLUMN suffix VARCHAR(100);
ALTER TABLE contacts ADD COLUMN phone_work VARCHAR(255);
ALTER TABLE contacts ADD COLUMN phone_work_2 VARCHAR(255);
ALTER TABLE contacts ADD COLUMN phone_mobile VARCHAR(255);
ALTER TABLE contacts ADD COLUMN phone_mobile_2 VARCHAR(255);
ALTER TABLE contacts ADD COLUMN phone_home VARCHAR(255);
ALTER TABLE contacts ADD COLUMN phone_fax_work VARCHAR(255);
ALTER TABLE contacts ADD COLUMN phone_fax_home VARCHAR(255);
ALTER TABLE contacts ADD COLUMN email2 VARCHAR(255);
ALTER TABLE contacts ADD COLUMN email3 VARCHAR(255);
ALTER TABLE contact_history ADD COLUMN change_type VARCHAR(50);
