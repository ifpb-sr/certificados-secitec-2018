require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'fileutils'
require 'yaml'
require 'byebug'
require 'csv'

OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'.freeze
APPLICATION_NAME = 'Google Sheets API Ruby Quickstart'.freeze
CREDENTIALS_PATH = 'credentials.json'.freeze
# The file token.yaml stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
TOKEN_PATH = 'token.yaml'.freeze
SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS_READONLY

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
#
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize
  client_id = Google::Auth::ClientId.from_file(CREDENTIALS_PATH)
  token_store = Google::Auth::Stores::FileTokenStore.new(file: TOKEN_PATH)
  authorizer = Google::Auth::UserAuthorizer.new(client_id, SCOPE, token_store)
  user_id = 'default'
  credentials = authorizer.get_credentials(user_id)
  if credentials.nil?
    url = authorizer.get_authorization_url(base_url: OOB_URI)
    puts 'Open the following URL in the browser and enter the ' \
         "resulting code after authorization:\n" + url
    code = gets
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI
    )
  end
  credentials
end


# atividades = ['SEGUNDA1930', 'TERCA0830', 'TERCA1000-1', 'TERCA1000-2', 'TERCA1000-3', 'TERCA1000-4', 'TERCA1330-1', 'TERCA1330-2', 'TERCA1330-3', 'TERCA1330-4', 'QUARTA0830', 'QUARTA1000-1', 'QUARTA1000-2', 'QUARTA1000-3', 'QUARTA1000-4', 'QUARTA1330-1', 'QUARTA1330-2', 'QUARTA1330-3', 'QUARTA1330-4', 'QUINTA0830', 'QUINTA1000-1', 'QUINTA1000-2', 'QUINTA1000-3', 'QUINTA1000-4', 'QUINTA1330-1', 'QUINTA1330-2', 'QUINTA1330-3', 'QUINTA1330-4', 'SEXTA0800']


atividades = YAML.load_file('atividades.yml')

desc 'Gera os csv para mesclar com os certificados'
task 'csv' do
# Initialize the API
  service = Google::Apis::SheetsV4::SheetsService.new
  service.client_options.application_name = APPLICATION_NAME
  service.authorization = authorize

  # Prints the names and majors of students in a sample spreadsheet:
  # https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
  spreadsheet_id = '1v1OpBmgd-ws5EbSQ9PahMKsfXMtOJWrAubtu5m2oZvg'
  range = 'TERCA1330_1'
  response = service.get_spreadsheet_values(spreadsheet_id, range)
  puts 'No data found.' if response.values.empty?
  response.values.each do |row|
    # Print columns A and E, which correspond to indices 0 and 4.
    puts "#{row[0]},#{row[1]}"
  end

end


# Palestras
{'SEGUNDA1930'=>'Ciência para a Redução das Desigualdades por meio da Educação Cooperativa'}.each do |horario,titulo|

  desc "Gera certificado da palestra realizada em #{horario}"
  task horario do
    sh "inkscape_merge -f modelos/certificado_ouvinte_de_palestra.svg -d modelos/certificado_ouvinte_de_palestra.csv -o certificados-gerados/certificado_ouvinte_de_palestra/#{horario}_%d.pdf"
  end



end

def gera_certificado(nome, modelo, titulo, data, chave)
  return if modelo == 'ApresentacaoOral'
  #byebug
  cvs_dir = "tmp/csv/#{nome}/"
  pdf_dir = "tmp/pdf/#{nome}/"
  FileUtils.mkdir_p(cvs_dir) unless File.directory?(cvs_dir)
  FileUtils.mkdir_p(pdf_dir) unless File.directory?(pdf_dir)
  csvfile = "tmp/csv/#{nome}/#{chave}.csv"
  pdffile = "tmp/pdf/#{nome}/#{chave}.pdf"
  modelofile = "modelos/digital/#{modelo}.svg"
  CSV.open(csvfile, "wb") do |csv|
    csv << ["nome", "titulo", "data"]
    csv << [nome, titulo, data]
  end

  sh "inkscape_merge -f #{modelofile} -d '#{csvfile}' -o '#{pdffile}'"

end


desc "Gera os certificados dos ouvintes a partir do google docs"
task :ouvintes do
  service = Google::Apis::SheetsV4::SheetsService.new
  service.client_options.application_name = APPLICATION_NAME
  service.authorization = authorize

  # Prints the names and majors of students in a sample spreadsheet:
  # https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
  spreadsheet_id = '1oo3zJuUMsoT8Mjde7ekgaGTTK2nyO2KBKkQWTmtQvg8'
  range = 'Página1'
  response = service.get_spreadsheet_values(spreadsheet_id, range)
  puts 'No data found.' if response.values.empty?
  presenca = {}
  response.values.each do |row|
    # Print columns A and E, which correspond to indices 0 and 4.
    puts "Gerando certificados do aluno: #{row[0]}"

    nome = row[0]
    presenca[nome] = 0

    atividades.each do |atividade|
      chave = atividade[0]
      titulo = atividade[1]['t']
      i = atividade[1]['i']
      modelo = atividade[1]['m']
      data = atividade[1]['d']
      presente = row[i]=="1"
      #puts "#{chave},#{presente}"


      if presente
        gera_certificado(nome, modelo, titulo, data, chave)
        presenca[nome] = presenca[nome] + 1 # Calcula presença
        registra_presenca(nome,chave,presenca)
      end

    end

    registra_justificativa(nome)
    #puts presenca
  end
end

# certificado fechine
