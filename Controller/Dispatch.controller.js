const { Pool } = require('pg');

const pool = new Pool({
  user: process.env.PGUSER,
  password: process.env.PGPASSWORD,
  host: process.env.PGHOST,
  port: process.env.PGPORT,
  database: process.env.PGDATABASE,
});

const roundVal = (v) => Math.round(parseFloat(v) || 0);

exports.getFilteredBookings = async (req, res) => {
  try {
    const { status } = req.query;
    const allowedStatuses = ['paid', 'packed', 'dispatched', 'delivered'];

    let query = `
      SELECT id, order_id, customer_name, district, state, status, mobile_number,
             total, products, address, customer_type,
             transport_name, lr_number, transport_contact,
             processing_date, dispatch_date, delivery_date,
             payment_method, transaction_id, amount_paid
      FROM public.bookings
      WHERE status = ANY($1)
    `;
    const params = [allowedStatuses];

    if (status) {
      const requestedStatuses = status.split(',').filter(s => allowedStatuses.includes(s));
      if (requestedStatuses.length > 0) {
        query = `
          SELECT id, order_id, customer_name, district, state, status, mobile_number,
                 total, products, address, customer_type,
                 transport_name, lr_number, transport_contact,
                 processing_date, dispatch_date, delivery_date,
                 payment_method, transaction_id, amount_paid
          FROM public.bookings
          WHERE status = ANY($1)
        `;
        params[0] = requestedStatuses;
      }
    }

    query += ` ORDER BY id DESC`;
    const result = await pool.query(query, params);
    res.status(200).json(result.rows);
  } catch (err) {
    console.error('Error fetching filtered bookings:', err);
    res.status(500).json({ message: 'Failed to fetch filtered bookings' });
  }
};

exports.getBookings = async (req, res) => {
  try {
    const { status, customerType } = req.query;

    let query = `
      SELECT id, order_id, customer_name, district, state, status, mobile_number,
             total, products, address, customer_type, created_at,
             transport_name, lr_number, transport_contact,
             processing_date, dispatch_date, delivery_date,
             payment_method, transaction_id, amount_paid
      FROM public.bookings
      WHERE status IN ('booked', 'paid', 'packed', 'dispatched', 'delivered')
    `;
    const params = [];
    let paramIndex = 1;

    if (status) {
      query += ` AND status = $${paramIndex++}`;
      params.push(status);
    }
    if (customerType) {
      query += ` AND customer_type = $${paramIndex++}`;
      params.push(customerType);
    }

    query += ` ORDER BY id DESC`;
    const result = await pool.query(query, params);
    res.status(200).json(result.rows);
  } catch (err) {
    console.error('Error fetching bookings:', err);
    res.status(500).json({ message: 'Failed to fetch bookings' });
  }
};

exports.updateBooking = async (req, res) => {
  try {
    const { order_id } = req.params;
    const { 
      products, net_rate, you_save, total, promo_discount, 
      additional_discount, status, 
      transportName, lrNumber, transportContact,   // New direct fields
      transport_details   // Keep for backward compatibility
    } = req.body;

    if (!order_id || !/^[a-zA-Z0-9-_]+$/.test(order_id))
      return res.status(400).json({ message: 'Invalid or missing Order ID', order_id });

    const bookingCheck = await pool.query('SELECT * FROM public.bookings WHERE order_id = $1', [order_id]);
    if (bookingCheck.rows.length === 0)
      return res.status(404).json({ message: 'Booking not found', order_id });

    const updateFields = [];
    const updateValues = [];
    let paramIndex = 1;

    // Existing updates (products, prices, etc.)
    if (products) {
      updateFields.push(`products = $${paramIndex++}`);
      updateValues.push(JSON.stringify(products));
    }
    if (net_rate !== undefined) {
      updateFields.push(`net_rate = $${paramIndex++}`);
      updateValues.push(roundVal(net_rate));
    }
    if (you_save !== undefined) {
      updateFields.push(`you_save = $${paramIndex++}`);
      updateValues.push(roundVal(you_save));
    }
    if (total !== undefined) {
      updateFields.push(`total = $${paramIndex++}`);
      updateValues.push(roundVal(total));
    }
    if (promo_discount !== undefined) {
      updateFields.push(`promo_discount = $${paramIndex++}`);
      updateValues.push(roundVal(promo_discount));
    }
    if (additional_discount !== undefined) {
      updateFields.push(`additional_discount = $${paramIndex++}`);
      updateValues.push(parseFloat(additional_discount) || 0);
    }
    if (status) {
      updateFields.push(`status = $${paramIndex++}`);
      updateValues.push(status);
    }

    // === TRANSPORT DETAILS - FIXED ===
    if (status === 'dispatched' || transportName || lrNumber || transport_details) {
      if (transportName || (transport_details && transport_details.transportName)) {
        updateFields.push(`transport_name = $${paramIndex++}`);
        updateValues.push(transportName || transport_details?.transportName || transport_details?.transport_name);
      }
      if (lrNumber || (transport_details && transport_details.lrNumber)) {
        updateFields.push(`lr_number = $${paramIndex++}`);
        updateValues.push(lrNumber || transport_details?.lrNumber || transport_details?.lr_number);
      }
      if (transportContact !== undefined || (transport_details && transport_details.transportContact)) {
        updateFields.push(`transport_contact = $${paramIndex++}`);
        updateValues.push(transportContact || transport_details?.transportContact || transport_details?.transport_contact || null);
      }
      if (status === 'dispatched') {
        updateFields.push(`dispatch_date = NOW()`);
      }
    }

    updateFields.push(`updated_at = NOW()`);

    if (updateFields.length === 1) {
      return res.status(400).json({ message: 'No fields to update', order_id });
    }

    const query = `
      UPDATE public.bookings
      SET ${updateFields.join(', ')}
      WHERE order_id = $${paramIndex}
      RETURNING id, order_id, status, transport_name, lr_number, transport_contact, dispatch_date
    `;
    updateValues.push(order_id);

    const result = await pool.query(query, updateValues);

    res.json({
      message: 'Booking updated successfully',
      data: result.rows[0]
    });

  } catch (err) {
    console.error(`Failed to update booking for order_id ${req.params.order_id}:`, err.message);
    res.status(500).json({ message: 'Failed to update booking', error: err.message });
  }
};

// Used by Dispatch.jsx — handles dispatched with transport details
// Used by Dispatch.jsx — handles dispatched with transport details
exports.updateFilteredBookingStatus = async (req, res) => {
  let client;
  try {
    const { id } = req.params;
    const { status, transportName, lrNumber, transportContact } = req.body;

    if (!id) return res.status(400).json({ message: 'Booking ID required' });

    client = await pool.connect();
    await client.query('BEGIN');

    // Update bookings table status
    const bookingUpdate = await client.query(`
      UPDATE public.bookings 
      SET status = $1, updated_at = NOW()
      WHERE id = $2
      RETURNING id, order_id
    `, [status, id]);

    if (bookingUpdate.rows.length === 0) {
      await client.query('ROLLBACK');
      return res.status(404).json({ message: 'Booking not found' });
    }

    const booking = bookingUpdate.rows[0];

    if (status === 'dispatched') {
      if (!transportName || !lrNumber) {
        await client.query('ROLLBACK');
        return res.status(400).json({ message: 'Transport Name and LR Number are required' });
      }

      // Insert into transport_details table
      await client.query(`
        INSERT INTO public.transport_details 
          (booking_id, order_id, transport_name, lr_number, transport_contact, created_at)
        VALUES ($1, $2, $3, $4, $5, NOW())
        ON CONFLICT (booking_id) DO UPDATE 
        SET transport_name = EXCLUDED.transport_name,
            lr_number = EXCLUDED.lr_number,
            transport_contact = EXCLUDED.transport_contact,
            updated_at = NOW()
      `, [id, booking.order_id, transportName, lrNumber, transportContact || null]);

      // Also update bookings table for quick access
      await client.query(`
        UPDATE public.bookings 
        SET transport_name = $1, 
            lr_number = $2, 
            transport_contact = $3,
            dispatch_date = NOW()
        WHERE id = $4
      `, [transportName, lrNumber, transportContact || null, id]);
    }

    await client.query('COMMIT');

    res.status(200).json({
      message: 'Status updated successfully',
      data: {
        id,
        status,
        transport_name: transportName,
        lr_number: lrNumber,
        transport_contact: transportContact
      }
    });

  } catch (err) {
    if (client) await client.query('ROLLBACK');
    console.error('Error updating status:', err);
    res.status(500).json({ message: 'Failed to update status', error: err.message });
  } finally {
    if (client) client.release();
  }
};  

exports.deleteBooking = async (req, res) => {
  try {
    const { order_id } = req.params;

    if (!order_id || !/^[a-zA-Z0-9-_]+$/.test(order_id)) {
      return res.status(400).json({ message: 'Invalid or missing Order ID' });
    }

    const result = await pool.query(
      'DELETE FROM public.bookings WHERE order_id = $1 RETURNING id, order_id',
      [order_id]
    );

    if (result.rows.length === 0) {
      return res.status(404).json({ message: 'Booking not found' });
    }

    res.status(200).json({ message: 'Booking deleted successfully', order_id });
  } catch (err) {
    console.error('Error deleting booking:', err);
    res.status(500).json({ message: 'Failed to delete booking' });
  }
};