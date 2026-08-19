module.exports = (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-store');
  res.status(200).json({
    supabaseUrl: process.env.VISIO_SUPABASE_URL || '',
    supabaseAnonKey: process.env.VISIO_SUPABASE_ANON_KEY || '',
    backendIaUrl: process.env.VISIO_IA_URL || ''
  });
};
