class EwsConnection
  
  # A separate connection class made it easy to mock the connection
  # in the exchange service tests. 
  
  def initialize(user, password, endpoint)
    @user, @password, @endpoint = user, password, endpoint
  end
  
  # Uses cURL to connect to the EWS server 2010 .
  # Passes the Soap XML document as the data.
  def connect(xml_doc)
  #  wsdl = `curl -u #{@user}:#{@password} -L #{@endpoint} -d "#{xml_doc.write}" -H "Content-Type:text/xml" --ntlm`
  #  Not using NTLM as there were issues, for time being only basic authentication is used
    connect_string = `curl -v -u #{@user}:#{@password} -L #{@endpoint} -d '#{xml_doc}' -H "Content-Type:text/xml" --ntlm`  
  end
  
end
