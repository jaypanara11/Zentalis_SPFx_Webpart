import * as React from 'react';
import { IZentalisRecognitionWallProps } from './IZentalisRecognitionWallProps';
import { IDateTimeFieldValue } from '@pnp/spfx-property-controls';
//import styles from './ZentalisRecognitionWall.module.scss';
export default class ZentalisRecognitionWall extends React.Component<IZentalisRecognitionWallProps, {}> {

  private formatDate(date: IDateTimeFieldValue | string | Date): string {
    const options: Intl.DateTimeFormatOptions = { month: 'short', day: 'numeric', year: 'numeric' };

    if (typeof date === 'string') {
      return new Date(date).toLocaleDateString('en-US', options);
    } else if (date instanceof Date) {
      return date.toLocaleDateString('en-US', options);
    } else if (date && (date as any).value) {
      // Assume 'value' holds the date string or timestamp
      const dateValue = (date as any).value;
      return new Date(dateValue).toLocaleDateString('en-US', options);
    }
    return ''; // Fallback if date is invalid
  }


  componentDidMount() {
    document.getElementById('View-All-Id')?.addEventListener('click', () => {
      const element = document.querySelector('.viewrs-slides_right') as HTMLElement;
      if (element) {
        element.style.display = element.style.display === 'none' ? 'block' : 'none';
      }
    });
  }
  
  public render(): React.ReactElement<IZentalisRecognitionWallProps> {
    const { selectedUsers } = this.props;
    const { selectedUsers1 } = this.props;
    const { selectedUsers2 } = this.props;
    const { selectedUsers3 } = this.props;
    const { selectedUsers4 } = this.props;
    const { selectedUsers5 } = this.props;
    const { selectedUsers6 } = this.props;
    const { selectedUsers7 } = this.props;
    const { selectedUsers8 } = this.props;
    const { selectedUsers9 } = this.props;
    const { selectedUsers10 } = this.props;
    const { selectedUsers11 } = this.props;
    
    return (
<div>
      <div className="container">
        <div className="recognition">
          <p id="p1">{this.props.Header}</p>
          <h1>{this.props.Title}</h1>
          <p id="p2">{this.props.Description}</p>
          <div className="footer_link">
            <a href={this.props.link}>{this.props.LinkText}</a>
            <img src={this.props.Icon} alt="" />
          </div>
        </div>
      </div>

  <div className="TilesView">
    <section className="views">
      <div className="viewrs-slides_left">
        <div className="reviewrs" style = {{ display: this.props.showCard1 ?'block':'none'}}>
          <div className="view-profile">
            <div className="view-img">            
                {selectedUsers && selectedUsers.map(user => ( 
                  <div key={user.id}>         
                    <img src={user.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                  </div>
                ))}            
                {selectedUsers1 && selectedUsers1.map(user1 => ( 
                  <div key={user1.id}> 
                    <img src={user1.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                  </div>
                ))}            
            </div>
            <div className="view-name">
              <div className="view_inside_name">
                <h4 className="first_name">
                  {selectedUsers && selectedUsers.map(user => ( 
                    <div key={user.id}>         
                      {user.fullName}
                    </div>
                  ))} 
                </h4>
                <p>recognized for being</p>
                <h4 className="last_name">{this.props.recognizedText1}</h4>
              </div>
              <div className="view_day">
                <p className="view_day_date">{this.formatDate(this.props.Date1)}</p>
                <p className="view_day_author">• by 
                  {selectedUsers1 && selectedUsers1.map(user1 => ( 
                    <div key={user1.id}>         
                      {user1.fullName}
                    </div>
                  ))}
                </p>
              </div>
            </div>
            <div className="view-icon">
              <img src={this.props.Icon1} alt=""/>
            </div>
          </div>
          <div className="view-text">
            <p>
              <span>"</span>{this.props.Description1}<span>"</span>
              </p>
          </div>
        </div>

        <div className="reviewrs" style = {{ display: this.props.showCard2 ?'block':'none'}}>
          <div className="view-profile">
            <div className="view-img">            
                {selectedUsers2 && selectedUsers2.map(user2 => ( 
                  <div key={user2.id}>         
                    <img src={user2.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                  </div>
                ))}            
                {selectedUsers3 && selectedUsers3.map(user3 => ( 
                  <div key={user3.id}> 
                    <img src={user3.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                  </div>
                ))}            
            </div>
            <div className="view-name">
              <div className="view_inside_name">
                <h4 className="first_name">
                  {selectedUsers2 && selectedUsers2.map(user2 => ( 
                    <div key={user2.id}>         
                      {user2.fullName}
                    </div>
                  ))} 
                </h4>
                <p>recognized for being</p>
                <h4 className="last_name">{this.props.recognizedText2}</h4>
              </div>
              <div className="view_day">
                <p className="view_day_date">{this.formatDate(this.props.Date2)}</p>
                <p className="view_day_author">• by 
                  {selectedUsers3 && selectedUsers3.map(user3 => ( 
                    <div key={user3.id}>         
                      {user3.fullName}
                    </div>
                  ))}
                </p>
              </div>
            </div>
            <div className="view-icon">
              <img src={this.props.Icon2} alt=""/>
            </div>
          </div>
          <div className="view-text">
            <p>
              <span>"</span>{this.props.Description2}<span>"</span>
              </p>
          </div>
        </div>   

        <div className="reviewrs" style = {{ display: this.props.showCard3 ?'block':'none'}}>
          <div className="view-profile">
            <div className="view-img">            
                {selectedUsers4 && selectedUsers4.map(user4 => ( 
                  <div key={user4.id}>         
                    <img src={user4.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                  </div>
                ))}            
                {selectedUsers5 && selectedUsers5.map(user5 => ( 
                  <div key={user5.id}> 
                    <img src={user5.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                  </div>
                ))}            
            </div>
            <div className="view-name">
              <div className="view_inside_name">
                <h4 className="first_name">
                  {selectedUsers4 && selectedUsers4.map(user4 => ( 
                    <div key={user4.id}>         
                      {user4.fullName}
                    </div>
                  ))} 
                </h4>
                <p>recognized for being</p>
                <h4 className="last_name">{this.props.recognizedText3}</h4>
              </div>
              <div className="view_day">
                <p className="view_day_date">{this.formatDate(this.props.Date3)}</p>
                <p className="view_day_author">• by 
                  {selectedUsers5 && selectedUsers5.map(user5 => ( 
                    <div key={user5.id}>         
                      {user5.fullName}
                    </div>
                  ))}
                </p>
              </div>
            </div>
            <div className="view-icon">
              <img src={this.props.Icon3} alt=""/>
            </div>
          </div>
          <div className="view-text">
            <p>
              <span>"</span>{this.props.Description3}<span>"</span>
              </p>
          </div>
        </div>
      </div>

      <span className="Mobile-View-Hide">
        <div className="viewrs-slides_left">
          <div className="reviewrs" style = {{ display: this.props.showCard1 ?'block':'none'}}>
            <div className="view-profile">
              <div className="view-img">            
                  {selectedUsers && selectedUsers.map(user => ( 
                    <div key={user.id}>         
                      <img src={user.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                    </div>
                  ))}            
                  {selectedUsers1 && selectedUsers1.map(user1 => ( 
                    <div key={user1.id}> 
                      <img src={user1.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                    </div>
                  ))}            
              </div>
              <div className="view-name">
                <div className="view_inside_name">
                  <h4 className="first_name">
                    {selectedUsers && selectedUsers.map(user => ( 
                      <div key={user.id}>         
                        {user.fullName}
                      </div>
                    ))} 
                  </h4>
                  <p>recognized for being</p>
                  <h4 className="last_name">{this.props.recognizedText1}</h4>
                </div>
                <div className="view_day">
                  <p className="view_day_date">{this.formatDate(this.props.Date1)}</p>
                  <p className="view_day_author">• by 
                    {selectedUsers1 && selectedUsers1.map(user1 => ( 
                      <div key={user1.id}>         
                        {user1.fullName}
                      </div>
                    ))}
                  </p>
                </div>
              </div>
              <div className="view-icon">
                <img src={this.props.Icon1} alt=""/>
              </div>
            </div>
            <div className="view-text">
              <p>
                <span>"</span>{this.props.Description1}<span>"</span>
                </p>
            </div>
          </div>

          <div className="reviewrs" style = {{ display: this.props.showCard2 ?'block':'none'}}>
            <div className="view-profile">
              <div className="view-img">            
                  {selectedUsers2 && selectedUsers2.map(user2 => ( 
                    <div key={user2.id}>         
                      <img src={user2.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                    </div>
                  ))}            
                  {selectedUsers3 && selectedUsers3.map(user3 => ( 
                    <div key={user3.id}> 
                      <img src={user3.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                    </div>
                  ))}            
              </div>
              <div className="view-name">
                <div className="view_inside_name">
                  <h4 className="first_name">
                    {selectedUsers2 && selectedUsers2.map(user2 => ( 
                      <div key={user2.id}>         
                        {user2.fullName}
                      </div>
                    ))} 
                  </h4>
                  <p>recognized for being</p>
                  <h4 className="last_name">{this.props.recognizedText2}</h4>
                </div>
                <div className="view_day">
                  <p className="view_day_date">{this.formatDate(this.props.Date2)}</p>
                  <p className="view_day_author">• by 
                    {selectedUsers3 && selectedUsers3.map(user3 => ( 
                      <div key={user3.id}>         
                        {user3.fullName}
                      </div>
                    ))}
                  </p>
                </div>
              </div>
              <div className="view-icon">
                <img src={this.props.Icon2} alt=""/>
              </div>
            </div>
            <div className="view-text">
              <p>
                <span>"</span>{this.props.Description2}<span>"</span>
                </p>
            </div>
          </div>   

          <div className="reviewrs" style = {{ display: this.props.showCard3 ?'block':'none'}}>
            <div className="view-profile">
              <div className="view-img">            
                  {selectedUsers4 && selectedUsers4.map(user4 => ( 
                    <div key={user4.id}>         
                      <img src={user4.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                    </div>
                  ))}            
                  {selectedUsers5 && selectedUsers5.map(user5 => ( 
                    <div key={user5.id}> 
                      <img src={user5.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                    </div>
                  ))}            
              </div>
              <div className="view-name">
                <div className="view_inside_name">
                  <h4 className="first_name">
                    {selectedUsers4 && selectedUsers4.map(user4 => ( 
                      <div key={user4.id}>         
                        {user4.fullName}
                      </div>
                    ))} 
                  </h4>
                  <p>recognized for being</p>
                  <h4 className="last_name">{this.props.recognizedText3}</h4>
                </div>
                <div className="view_day">
                  <p className="view_day_date">{this.formatDate(this.props.Date3)}</p>
                  <p className="view_day_author">• by 
                    {selectedUsers5 && selectedUsers5.map(user5 => ( 
                      <div key={user5.id}>         
                        {user5.fullName}
                      </div>
                    ))}
                  </p>
                </div>
              </div>
              <div className="view-icon">
                <img src={this.props.Icon3} alt=""/>
              </div>
            </div>
            <div className="view-text">
              <p>
                <span>"</span>{this.props.Description3}<span>"</span>
                </p>
            </div>
          </div>
        </div>
      </span>
    </section>

    <section className="views">
      <div className="viewrs-slides_right">
        <div className="reviewrs" style = {{ display: this.props.showCard4 ?'block':'none'}}>
          <div className="view-profile">
            <div className="view-img">            
                {selectedUsers6 && selectedUsers6.map(user6 => ( 
                  <div key={user6.id}>         
                    <img src={user6.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                  </div>
                ))}            
                {selectedUsers7 && selectedUsers7.map(user7 => ( 
                  <div key={user7.id}> 
                    <img src={user7.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                  </div>
                ))}            
            </div>
            <div className="view-name">
              <div className="view_inside_name">
                <h4 className="first_name">
                  {selectedUsers6 && selectedUsers6.map(user6 => ( 
                    <div key={user6.id}>         
                      {user6.fullName}
                    </div>
                  ))} 
                </h4>
                <p>recognized for being</p>
                <h4 className="last_name">{this.props.recognizedText4}</h4>
              </div>
              <div className="view_day">
                <p className="view_day_date">{this.formatDate(this.props.Date4)}</p>
                <p className="view_day_author">• by 
                  {selectedUsers7 && selectedUsers7.map(user7 => ( 
                    <div key={user7.id}>         
                      {user7.fullName}
                    </div>
                  ))}
                </p>
              </div>
            </div>
            <div className="view-icon">
              <img src={this.props.Icon4} alt=""/>
            </div>
          </div>
          <div className="view-text">
            <p>
              <span>"</span>{this.props.Description4}<span>"</span>
              </p>
          </div>
        </div>

        <div className="reviewrs" style = {{ display: this.props.showCard5 ?'block':'none'}}>
          <div className="view-profile">
            <div className="view-img">            
                {selectedUsers8 && selectedUsers8.map(user8 => ( 
                  <div key={user8.id}>         
                    <img src={user8.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                  </div>
                ))}            
                {selectedUsers9 && selectedUsers9.map(user9 => ( 
                  <div key={user9.id}> 
                    <img src={user9.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                  </div>
                ))}            
            </div>
            <div className="view-name">
              <div className="view_inside_name">
                <h4 className="first_name">
                  {selectedUsers8 && selectedUsers8.map(user8 => ( 
                    <div key={user8.id}>         
                      {user8.fullName}
                    </div>
                  ))} 
                </h4>
                <p>recognized for being</p>
                <h4 className="last_name">{this.props.recognizedText5}</h4>
              </div>
              <div className="view_day">
                <p className="view_day_date">{this.formatDate(this.props.Date5)}</p>
                <p className="view_day_author">• by 
                  {selectedUsers9 && selectedUsers9.map(user9 => ( 
                    <div key={user9.id}>         
                      {user9.fullName}
                    </div>
                  ))}
                </p>
              </div>
            </div>
            <div className="view-icon">
              <img src={this.props.Icon5} alt=""/>
            </div>
          </div>
          <div className="view-text">
            <p>
              <span>"</span>{this.props.Description5}<span>"</span>
              </p>
          </div>
        </div>   

        <div className="reviewrs" style = {{ display: this.props.showCard6 ?'block':'none'}}>
          <div className="view-profile">
            <div className="view-img">            
                {selectedUsers10 && selectedUsers10.map(user10 => ( 
                  <div key={user10.id}>         
                    <img src={user10.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                  </div>
                ))}            
                {selectedUsers11 && selectedUsers11.map(user11 => ( 
                  <div key={user11.id}> 
                    <img src={user11.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                  </div>
                ))}            
            </div>
            <div className="view-name">
              <div className="view_inside_name">
                <h4 className="first_name">
                  {selectedUsers10 && selectedUsers10.map(user10 => ( 
                    <div key={user10.id}>         
                      {user10.fullName}
                    </div>
                  ))} 
                </h4>
                <p>recognized for being</p>
                <h4 className="last_name">{this.props.recognizedText6}</h4>
              </div>
              <div className="view_day">
                <p className="view_day_date">{this.formatDate(this.props.Date6)}</p>
                <p className="view_day_author">• by 
                  {selectedUsers11 && selectedUsers11.map(user11 => ( 
                    <div key={user11.id}>         
                      {user11.fullName}
                    </div>
                  ))}
                </p>
              </div>
            </div>
            <div className="view-icon">
              <img src={this.props.Icon6} alt=""/>
            </div>
          </div>
          <div className="view-text">
            <p>
              <span>"</span>{this.props.Description6}<span>"</span>
              </p>
          </div>
        </div>
      </div>

      <span className="Mobile-View-Hide">
        <div className="viewrs-slides_right">
          <div className="reviewrs" style = {{ display: this.props.showCard4 ?'block':'none'}}>
            <div className="view-profile">
              <div className="view-img">            
                  {selectedUsers6 && selectedUsers6.map(user6 => ( 
                    <div key={user6.id}>         
                      <img src={user6.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                    </div>
                  ))}            
                  {selectedUsers7 && selectedUsers7.map(user7 => ( 
                    <div key={user7.id}> 
                      <img src={user7.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                    </div>
                  ))}            
              </div>
              <div className="view-name">
                <div className="view_inside_name">
                  <h4 className="first_name">
                    {selectedUsers6 && selectedUsers6.map(user6 => ( 
                      <div key={user6.id}>         
                        {user6.fullName}
                      </div>
                    ))} 
                  </h4>
                  <p>recognized for being</p>
                  <h4 className="last_name">{this.props.recognizedText4}</h4>
                </div>
                <div className="view_day">
                  <p className="view_day_date">{this.formatDate(this.props.Date4)}</p>
                  <p className="view_day_author">• by 
                    {selectedUsers7 && selectedUsers7.map(user7 => ( 
                      <div key={user7.id}>         
                        {user7.fullName}
                      </div>
                    ))}
                  </p>
                </div>
              </div>
              <div className="view-icon">
                <img src={this.props.Icon4} alt=""/>
              </div>
            </div>
            <div className="view-text">
              <p>
                <span>"</span>{this.props.Description4}<span>"</span>
                </p>
            </div>
          </div>

          <div className="reviewrs" style = {{ display: this.props.showCard5 ?'block':'none'}}>
            <div className="view-profile">
              <div className="view-img">            
                  {selectedUsers8 && selectedUsers8.map(user8 => ( 
                    <div key={user8.id}>         
                      <img src={user8.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                    </div>
                  ))}            
                  {selectedUsers9 && selectedUsers9.map(user9 => ( 
                    <div key={user9.id}> 
                      <img src={user9.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                    </div>
                  ))}            
              </div>
              <div className="view-name">
                <div className="view_inside_name">
                  <h4 className="first_name">
                    {selectedUsers8 && selectedUsers8.map(user8 => ( 
                      <div key={user8.id}>         
                        {user8.fullName}
                      </div>
                    ))} 
                  </h4>
                  <p>recognized for being</p>
                  <h4 className="last_name">{this.props.recognizedText5}</h4>
                </div>
                <div className="view_day">
                  <p className="view_day_date">{this.formatDate(this.props.Date5)}</p>
                  <p className="view_day_author">• by 
                    {selectedUsers9 && selectedUsers9.map(user9 => ( 
                      <div key={user9.id}>         
                        {user9.fullName}
                      </div>
                    ))}
                  </p>
                </div>
              </div>
              <div className="view-icon">
                <img src={this.props.Icon5} alt=""/>
              </div>
            </div>
            <div className="view-text">
              <p>
                <span>"</span>{this.props.Description5}<span>"</span>
                </p>
            </div>
          </div>   

          <div className="reviewrs" style = {{ display: this.props.showCard6 ?'block':'none'}}>
            <div className="view-profile">
              <div className="view-img">            
                  {selectedUsers10 && selectedUsers10.map(user10 => ( 
                    <div key={user10.id}>         
                      <img src={user10.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-one"/>
                    </div>
                  ))}            
                  {selectedUsers11 && selectedUsers11.map(user11 => ( 
                    <div key={user11.id}> 
                      <img src={user11.imageUrl} alt="" style={{ width: '32px', height: '32px', borderRadius: '50%'}} className="img-two"/>
                    </div>
                  ))}            
              </div>
              <div className="view-name">
                <div className="view_inside_name">
                  <h4 className="first_name">
                    {selectedUsers10 && selectedUsers10.map(user10 => ( 
                      <div key={user10.id}>         
                        {user10.fullName}
                      </div>
                    ))} 
                  </h4>
                  <p>recognized for being</p>
                  <h4 className="last_name">{this.props.recognizedText6}</h4>
                </div>
                <div className="view_day">
                  <p className="view_day_date">{this.formatDate(this.props.Date6)}</p>
                  <p className="view_day_author">• by 
                    {selectedUsers11 && selectedUsers11.map(user11 => ( 
                      <div key={user11.id}>         
                        {user11.fullName}
                      </div>
                    ))}
                  </p>
                </div>
              </div>
              <div className="view-icon">
                <img src={this.props.Icon6} alt=""/>
              </div>
            </div>
            <div className="view-text">
              <p>
                <span>"</span>{this.props.Description6}<span>"</span>
                </p>
            </div>
          </div>
        </div> 
      </span>
    </section>  
  </div>        
  <button className="View-All-Btn" id="View-All-Id" type="button">View All</button>
</div>   
 
    );   
  }  
}
